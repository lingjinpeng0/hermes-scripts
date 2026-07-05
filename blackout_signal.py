import sys, os

SIGNAL = r'C:\Users\Rei\.hermes\blackout_signal.tmp'

if __name__ == '__main__':
    cmd = sys.argv[1] if len(sys.argv) > 1 else ''
    if cmd in ('on', 'off'):
        with open(SIGNAL, 'w') as f:
            f.write(cmd)
        print(f'✅ 已发送{cmd}信号')
    else:
        print('用法: on / off')
