import pdf_gen_travel
from datetime import datetime
import traceback

record = {
    'form_date': '2023-10-25',
    'traveler_name': 'Test User',
    'plan_code': 'PLAN123',
    'purpose_desc': 'Visiting client site for an annual checkup.',
    'travel_route': 'Taipei - Hsinchu roundtrip',
    'start_time': '2023-10-26T08:00',
    'end_time': '2023-10-26T18:00',
    'travel_days': 1,
    'is_hsr': True,
    'estimated_cost': 1500,
    'handler_name': 'Handler A',
    'project_manager_name': 'PM B',
    'dept_manager_name': 'Manager C',
    'accountant_name': 'Accountant D',
    'expense_rows': '[{"type": "交通", "desc": "高鐵", "amount": 1500}]'
}
try:
    pdf_bytes = pdf_gen_travel.build_pdf_bytes(record, [])
    with open('test_travel_voucher.pdf', 'wb') as f:
        f.write(pdf_bytes)
    print('PDF generated successfully!')
except Exception:
    traceback.print_exc()
