"""Rebuild Hermes state.db, recovering as much data as possible despite b-tree corruption."""
import sqlite3, os, shutil, traceback

SRC = r'C:\Users\Rei\AppData\Local\hermes\state.db.recover'
DST = r'C:\Users\Rei\AppData\Local\hermes\state.db.rebuilt2'

srcc = sqlite3.connect(SRC)
srcc.text_factory = str

# Get column info
msg_cols_info = srcc.execute('PRAGMA table_info(messages);').fetchall()
msg_col_names = [c[1] for c in msg_cols_info]
print('MSG columns:', msg_col_names)

# Get all session IDs from messages (with error handling for the index)
try:
    sids = srcc.execute('SELECT DISTINCT session_id FROM messages;').fetchall()
    sids = [s[0] for s in sids]
except Exception:
    # If DISTINCT on session_id fails, try scanning raw messages
    print('DISTINCT failed, scanning by LIMIT/OFFSET...')
    sids = set()
    offset = 0
    while True:
        try:
            rows = srcc.execute('SELECT session_id FROM messages LIMIT 100 OFFSET ?', (offset,)).fetchall()
            if not rows:
                break
            for r in rows:
                sids.add(r[0])
            offset += 100
        except Exception:
            break
    sids = list(sids)

print('All session IDs:', len(sids))

# Build new DB
if os.path.exists(DST):
    os.remove(DST)
dstc = sqlite3.connect(DST)

# Create tables
table_schemas = srcc.execute(
    "SELECT sql FROM sqlite_master WHERE type='table' AND name NOT IN ('sqlite_sequence') AND name NOT LIKE '%fts%' ORDER BY rowid;"
).fetchall()
for (sql,) in table_schemas:
    dstc.execute(sql)

dstc.execute('PRAGMA journal_mode=WAL;')

# Copy sessions
sess_cols = [d[1] for d in srcc.execute('PRAGMA table_info(sessions);').fetchall()]
ph = ','.join(['?'] * len(sess_cols))
rows = srcc.execute('SELECT * FROM sessions ORDER BY rowid;').fetchall()
dstc.executemany('INSERT INTO sessions VALUES(' + ph + ')', rows)
print('sessions:', len(rows))

# Copy state_meta
meta_cols = [d[1] for d in srcc.execute('PRAGMA table_info(state_meta);').fetchall()]
ph = ','.join(['?'] * len(meta_cols))
rows = srcc.execute('SELECT * FROM state_meta;').fetchall()
dstc.executemany('INSERT INTO state_meta VALUES(' + ph + ')', rows)
print('state_meta:', len(rows))

# Copy sqlite_sequence
rows = srcc.execute('SELECT * FROM sqlite_sequence;').fetchall()
for r in rows:
    dstc.execute('INSERT OR REPLACE INTO sqlite_sequence VALUES(?,?)', r)
print('sqlite_sequence:', len(rows))

# Copy compression_locks
lock_cols = [d[1] for d in srcc.execute('PRAGMA table_info(compression_locks);').fetchall()]
if lock_cols:
    ph = ','.join(['?'] * len(lock_cols))
    rows = srcc.execute('SELECT * FROM compression_locks;').fetchall()
    dstc.executemany('INSERT INTO compression_locks VALUES(' + ph + ')', rows)
print('compression_locks:', len(rows))

# Copy schema_version
sv_cols = [d[1] for d in srcc.execute('PRAGMA table_info(schema_version);').fetchall()]
ph = ','.join(['?'] * len(sv_cols))
rows = srcc.execute('SELECT * FROM schema_version;').fetchall()
for r in rows:
    dstc.execute('INSERT OR REPLACE INTO schema_version VALUES(' + ph + ')', r)
print('schema_version:', len(rows))

# Copy messages - with per-session error handling
ph = ','.join(['?'] * len(msg_col_names))
total = 0
lost = 0
for sid in sids:
    try:
        rows = srcc.execute('SELECT * FROM messages WHERE session_id = ? ORDER BY rowid;', (sid,)).fetchall()
        if rows:
            dstc.executemany('INSERT INTO messages VALUES(' + ph + ')', rows)
            total += len(rows)
    except Exception as e:
        # Try without ORDER BY
        try:
            rows = srcc.execute('SELECT * FROM messages WHERE session_id = ?;', (sid,)).fetchall()
            if rows:
                dstc.executemany('INSERT INTO messages VALUES(' + ph + ')', rows)
                total += len(rows)
                print('  Partially recovered:', sid[:30], '(' + str(len(rows)) + ')')
        except Exception as e2:
            print('  LOST session:', sid[:30], '-', str(e2)[:60])
            lost += 1

print('messages:', total, '(lost sessions:', lost, ')')

# Create FTS tables
fts_sqls = srcc.execute(
    "SELECT name, sql FROM sqlite_master WHERE type='table' AND name LIKE '%fts%' ORDER BY rowid;"
).fetchall()
for name, sql in fts_sqls:
    try:
        dstc.execute(sql)
        print('  FTS:', name)
    except Exception as e:
        print('  FTS error:', name, str(e)[:60])

# Rebuild FTS indexes
try:
    dstc.execute("INSERT INTO messages_fts(messages_fts) VALUES('rebuild')")
    print('FTS rebuilt')
except Exception as e:
    print('FTS rebuild error:', str(e)[:100])

try:
    dstc.execute("INSERT INTO messages_fts_trigram(messages_fts_trigram) VALUES('rebuild')")
    print('FTS trigram rebuilt')
except Exception as e:
    print('FTS trigram error:', str(e)[:100])

dstc.commit()
dstc.close()
srcc.close()

# Final verification
vc = sqlite3.connect(DST)
cur = vc.execute('PRAGMA integrity_check;')
result = cur.fetchone()[0]
if result == 'ok':
    print('\n=== INTEGRITY: OK ===')
else:
    print('\n=== INTEGRITY:', len(result.splitlines()), 'issues ===')
    for line in result.splitlines()[:5]:
        print(' ', line)

cur = vc.execute('SELECT count(*) FROM sessions;')
print('sessions:', cur.fetchone()[0])
cur = vc.execute('SELECT count(*) FROM messages;')
print('messages:', cur.fetchone()[0])
# FTS check
try:
    cur = vc.execute("SELECT count(*) FROM messages_fts WHERE messages_fts MATCH 'test';")
    print('FTS works:', cur.fetchone()[0])
except Exception as e:
    print('FTS query failed:', str(e)[:100])
vc.close()

print('\nFile:', DST, '-', os.path.getsize(DST) / 1024 / 1024, 'MB')
