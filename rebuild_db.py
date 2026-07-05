import sqlite3, os, sys

TMP = r'C:\Users\Rei\tmpfinal.db'
NEW = r'C:\Users\Rei\AppData\Local\hermes\state_clean.db'

# 先获取所有存在的 ID
src = sqlite3.connect(TMP)
src.execute('PRAGMA recovery_mode = 1;')
cur = src.execute('SELECT COUNT(*) FROM messages;')
total = cur.fetchone()[0]
cur = src.execute('SELECT MIN(id), MAX(id) FROM messages;')
min_id, max_id = cur.fetchone()
print(f'Expected: {total} rows, ids {min_id}-{max_id}')
src.close()

# 分 chunk 读，每个 chunk 新连接+recovery
all_rows = []
lost = 0
chunk = 300

for start in range(min_id, max_id + 1, chunk):
    end = min(start + chunk - 1, max_id)
    src = sqlite3.connect(TMP)
    src.execute('PRAGMA recovery_mode = 1;')
    
    try:
        cur = src.execute('''SELECT id, session_id, role, content, tool_call_id,
            tool_calls, tool_name, timestamp, token_count, finish_reason,
            reasoning, reasoning_content, reasoning_details,
            codex_reasoning_items, codex_message_items, platform_message_id,
            observed, active, compacted
            FROM messages WHERE id >= ? AND id <= ? ORDER BY id;''', (start, end))
        rows = cur.fetchall()
        all_rows.extend(rows)
        src.close()
        if start % 1500 == 1:
            print(f'  [{start}-{end}] chunk OK, total: {len(all_rows)}')
        continue
    except:
        pass  # chunk failed
    
    # Chunk failed - try individual
    for rid in range(start, end + 1):
        try:
            cur = src.execute('''SELECT id, session_id, role, content, tool_call_id,
                tool_calls, tool_name, timestamp, token_count, finish_reason,
                reasoning, reasoning_content, reasoning_details,
                codex_reasoning_items, codex_message_items, platform_message_id,
                observed, active, compacted
                FROM messages WHERE id = ?;''', (rid,))
            row = cur.fetchone()
            if row:
                all_rows.append(row)
            else:
                lost += 1
        except:
            lost += 1
    src.close()
    
    if start % 1500 == 1:
        print(f'  [{start}-{end}] individual mode, total: {len(all_rows)}, lost: {lost}')

print(f'\nRecovered: {len(all_rows)} / {total}, lost: {lost}')

if len(all_rows) == 0:
    print('NO DATA! Aborting.')
    sys.exit(1)

print(f'First: id={all_rows[0][0]}')
print(f'Last: id={all_rows[-1][0]}')

# === Build clean DB ===
print('\n=== Building clean database ===')
new = sqlite3.connect(NEW)
new.execute('PRAGMA journal_mode=WAL;')

# Drop if exists for clean rebuild
new.execute('DROP TABLE IF EXISTS sessions;')
new.execute('DROP TABLE IF EXISTS messages;')
new.execute('DROP TABLE IF EXISTS state_meta;')
new.execute('DROP TABLE IF EXISTS schema_version;')
new.execute('DROP TABLE IF EXISTS compression_locks;')

new.execute('''CREATE TABLE sessions (
    id TEXT PRIMARY KEY, source TEXT NOT NULL, user_id TEXT,
    model TEXT, model_config TEXT, system_prompt TEXT,
    parent_session_id TEXT, started_at REAL NOT NULL, ended_at REAL,
    end_reason TEXT, message_count INTEGER DEFAULT 0,
    tool_call_count INTEGER DEFAULT 0, input_tokens INTEGER DEFAULT 0,
    output_tokens INTEGER DEFAULT 0, cache_read_tokens INTEGER DEFAULT 0,
    cache_write_tokens INTEGER DEFAULT 0, reasoning_tokens INTEGER DEFAULT 0,
    cwd TEXT, billing_provider TEXT, billing_base_url TEXT,
    billing_mode TEXT, estimated_cost_usd REAL, actual_cost_usd REAL,
    cost_status TEXT, cost_source TEXT, pricing_version TEXT, title TEXT,
    api_call_count INTEGER DEFAULT 0, handoff_state TEXT,
    handoff_platform TEXT, handoff_error TEXT,
    rewind_count INTEGER NOT NULL DEFAULT 0,
    archived INTEGER NOT NULL DEFAULT 0,
    FOREIGN KEY (parent_session_id) REFERENCES sessions(id)
);''')

new.execute('''CREATE TABLE messages (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    session_id TEXT NOT NULL REFERENCES sessions(id),
    role TEXT NOT NULL, content TEXT,
    tool_call_id TEXT, tool_calls TEXT, tool_name TEXT,
    timestamp REAL NOT NULL, token_count INTEGER,
    finish_reason TEXT, reasoning TEXT,
    reasoning_content TEXT, reasoning_details TEXT,
    codex_reasoning_items TEXT, codex_message_items TEXT,
    platform_message_id TEXT, observed INTEGER DEFAULT 0,
    active INTEGER NOT NULL DEFAULT 1,
    "compacted" INTEGER NOT NULL DEFAULT 0
);''')

new.execute('CREATE TABLE state_meta (key TEXT PRIMARY KEY, value TEXT);')
new.execute('CREATE TABLE schema_version (version INTEGER NOT NULL);')
# sqlite_sequence is auto-created by SQLite for AUTOINCREMENT - skip manual creation
new.execute('CREATE TABLE compression_locks (session_id TEXT PRIMARY KEY, holder TEXT NOT NULL, acquired_at REAL NOT NULL, expires_at REAL NOT NULL);')

# Read sessions etc from a separate fresh connection
src2 = sqlite3.connect(TMP)
src2.execute('PRAGMA recovery_mode = 1;')
cur = src2.execute('SELECT * FROM sessions;')
sess_rows = cur.fetchall()
print(f'sessions: {len(sess_rows)}')
cur = src2.execute('SELECT * FROM state_meta;')
meta_rows = cur.fetchall()
cur = src2.execute('SELECT * FROM schema_version;')
sv_rows = cur.fetchall()
cur = src2.execute('SELECT * FROM sqlite_sequence;')
sq_rows = cur.fetchall()
try:
    cur = src2.execute('SELECT * FROM compression_locks;')
    cl_rows = cur.fetchall()
except:
    cl_rows = []
src2.close()

# Insert base data
sess_cols = ['id','source','user_id','model','model_config','system_prompt',
    'parent_session_id','started_at','ended_at','end_reason','message_count',
    'tool_call_count','input_tokens','output_tokens','cache_read_tokens',
    'cache_write_tokens','reasoning_tokens','cwd','billing_provider',
    'billing_base_url','billing_mode','estimated_cost_usd','actual_cost_usd',
    'cost_status','cost_source','pricing_version','title','api_call_count',
    'handoff_state','handoff_platform','handoff_error','rewind_count','archived']
ph = ','.join(['?' for _ in sess_cols])
new.executemany(f'INSERT INTO sessions VALUES ({ph});', sess_rows)

msg_cols = ['id','session_id','role','content','tool_call_id','tool_calls',
    'tool_name','timestamp','token_count','finish_reason','reasoning',
    'reasoning_content','reasoning_details','codex_reasoning_items',
    'codex_message_items','platform_message_id','observed','active','compacted']
ph = ','.join(['?' for _ in msg_cols])
new.executemany(f'INSERT INTO messages VALUES ({ph});', all_rows)

ph = ','.join(['?' for _ in ['key','value']])
new.executemany(f'INSERT INTO state_meta VALUES ({ph});', meta_rows)
new.execute('INSERT INTO schema_version VALUES (?);', sv_rows[0])
if sq_rows:
    new.execute('INSERT INTO sqlite_sequence VALUES (?, ?);', sq_rows[0])
if cl_rows:
    ph = ','.join(['?' for _ in ['session_id','holder','acquired_at','expires_at']])
    new.executemany(f'INSERT INTO compression_locks VALUES ({ph});', cl_rows)

print('Base tables populated.')

# FTS
print('Creating FTS indexes...')
new.execute('CREATE VIRTUAL TABLE messages_fts USING fts5(content);')
new.execute('CREATE VIRTUAL TABLE messages_fts_trigram USING fts5(content, tokenize="trigram");')

cur = new.execute('SELECT id, content FROM messages WHERE content IS NOT NULL AND content != "";')
rows = cur.fetchall()
new.executemany('INSERT INTO messages_fts(rowid, content) VALUES (?, ?);', rows)
new.executemany('INSERT INTO messages_fts_trigram(rowid, content) VALUES (?, ?);', rows)
print(f'FTS populated: {len(rows)}')

new.commit()

# Verify
print()
print('=== Verification ===')
cur = new.execute('PRAGMA integrity_check;');
print(f'Integrity: {cur.fetchone()[0]}')

for tbl in ['sessions', 'messages', 'messages_fts', 'messages_fts_trigram']:
    cur = new.execute(f'SELECT COUNT(*) FROM "{tbl}";')
    print(f'{tbl}: {cur.fetchone()[0]}')

# FTS search test
cur = new.execute("SELECT rowid FROM messages_fts WHERE messages_fts MATCH 'test' LIMIT 5;")
print(f'FTS search ("test"): {len(cur.fetchall())}')

cur = new.execute("SELECT rowid FROM messages_fts_trigram WHERE messages_fts_trigram MATCH 'test' LIMIT 5;")
print(f'Trigram search ("test"): {len(cur.fetchall())}')

new.close()

old_sz = os.path.getsize(r'C:\Users\Rei\AppData\Local\hermes\state.db')
new_sz = os.path.getsize(NEW)
print(f'\nOld: {old_sz/1024/1024:.0f}MB → New: {new_sz/1024/1024:.0f}MB')
print(f'Data loss: {total - len(all_rows)} rows')
print('✅ state_clean.db ready!')
