from zoneinfo import ZoneInfo
from datetime import datetime

ist = datetime(2026, 3, 27, 8, 0, tzinfo=ZoneInfo('Asia/Kolkata'))
cst = ist.astimezone(ZoneInfo('Asia/Shanghai'))
print(f'IST 08:00 = CST {cst.hour}:{cst.minute:02d}')

ist_end = datetime(2026, 3, 27, 17, 0, tzinfo=ZoneInfo('Asia/Kolkata'))
cst_end = ist_end.astimezone(ZoneInfo('Asia/Shanghai'))
print(f'IST 17:00 = CST {cst_end.hour}:{cst_end.minute:02d}')

# Made with Bob
