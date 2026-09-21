import json
import os
from pathlib import Path

os.environ.setdefault("LAB_ADMIN_USERNAME", "admin")
os.environ.setdefault("LAB_ADMIN_PASSWORD", "TestPassword123!")

from app import BORROWED_PATH, REQUESTS_PATH, app, save_json

save_json(REQUESTS_PATH, [])
save_json(BORROWED_PATH, [])

client = app.test_client()

index = client.get('/')
assert index.status_code == 200, index.status_code
assert 'Lab Inventory Tracker' in index.get_data(as_text=True)
assert 'Borrowed Parts' not in index.get_data(as_text=True)

blocked_admin = client.get('/admin')
assert blocked_admin.status_code == 302, blocked_admin.status_code
assert '/login' in blocked_admin.headers.get('Location', '')

invalid_login = client.post('/login', data={'username': 'wrong', 'password': 'bad'})
assert invalid_login.status_code == 401, invalid_login.status_code

login = client.post(
    '/login',
    data={'username': os.environ['LAB_ADMIN_USERNAME'], 'password': os.environ['LAB_ADMIN_PASSWORD']},
)
assert login.status_code == 302, login.status_code
assert '/admin' in login.headers.get('Location', '')

before_invalid_requests = len(json.loads(Path('borrow_requests.json').read_text(encoding='utf-8')))
invalid_past_date = client.post(
    '/borrow',
    data={
        'item_name': 'Ethernet Cables',
        'quantity': '1',
        'borrower_name': 'Past Date User',
        'borrower_id': '777',
        'purpose': 'Testing',
        'return_date': '2020-01-01',
    },
)
assert invalid_past_date.status_code == 302, invalid_past_date.status_code
assert len(json.loads(Path('borrow_requests.json').read_text(encoding='utf-8'))) == before_invalid_requests

before_full_amount = len(json.loads(Path('borrow_requests.json').read_text(encoding='utf-8')))
equal_amount_response = client.post(
    '/borrow',
    data={
        'item_name': 'Ethernet Cables',
        'quantity': '10',
        'borrower_name': 'Full Amount User',
        'borrower_id': '777',
        'borrower_email': 'test@example.com',
        'purpose': 'Testing max quantity',
        'return_date': '2030-12-31',
    },
)
assert equal_amount_response.status_code in (200, 302), equal_amount_response.status_code
assert len(json.loads(Path('borrow_requests.json').read_text(encoding='utf-8'))) == before_full_amount + 1

requests = json.loads(Path('borrow_requests.json').read_text(encoding='utf-8'))
last_request = requests[-1]
assert last_request['item_name'] == 'Ethernet Cables'
assert last_request['status'] == 'pending'
assert last_request['qty'] == 10
assert last_request['borrower_email'] == 'test@example.com'

admin_page = client.get('/admin')
assert admin_page.status_code == 200, admin_page.status_code
assert 'Full Amount User' in admin_page.get_data(as_text=True)

approval = client.post(f"/admin/approve/{last_request['id']}")
assert approval.status_code == 302, approval.status_code

save_json(
    REQUESTS_PATH,
    [
        {
            'id': 'limit-approved',
            'item_name': 'Ethernet Cables',
            'qty': 8,
            'borrower_name': 'Approved User',
            'borrower_id': '8',
            'purpose': 'Already approved',
            'return_date': '2030-12-31',
            'location': 'C1 / Ethernet Cables',
            'status': 'approved',
            'submitted_at': '2026-09-21 09:00:00',
            'approved_at': '2026-09-21 09:00:00',
        },
        {
            'id': 'limit-pending',
            'item_name': 'Ethernet Cables',
            'qty': 3,
            'borrower_name': 'Over Limit User',
            'borrower_id': '999',
            'purpose': 'Should be rejected',
            'return_date': '2030-01-01',
            'location': 'C1 / Ethernet Cables',
            'status': 'pending',
            'submitted_at': '2026-09-21 09:05:00',
        },
    ],
)
reject_over_limit = client.post('/admin/approve/limit-pending')
assert reject_over_limit.status_code == 302, reject_over_limit.status_code
updated_requests = json.loads(Path('borrow_requests.json').read_text(encoding='utf-8'))
matching = next((req for req in updated_requests if req.get('id') == 'limit-pending'), None)
assert matching is not None and matching['status'] == 'pending', matching

save_json(
    REQUESTS_PATH,
    [
        {
            'id': 'final-approved',
            'item_name': 'Ethernet Cables',
            'qty': 1,
            'borrower_name': 'Test User',
            'borrower_id': '123',
            'purpose': 'Testing',
            'return_date': '2026-10-15',
            'location': 'C1 / Ethernet Cables',
            'status': 'approved',
            'submitted_at': '2026-09-21 09:10:00',
            'approved_at': '2026-09-21 09:10:00',
        }
    ],
)
approved_admin_page = client.get('/admin')
assert 'Test User' in approved_admin_page.get_data(as_text=True)
assert 'Approved Borrow Records' in approved_admin_page.get_data(as_text=True)

borrowed_admin_page = client.get('/admin/borrowed')
assert borrowed_admin_page.status_code == 200, borrowed_admin_page.status_code
assert 'Borrowed Items' in borrowed_admin_page.get_data(as_text=True)

cancel_response = client.post('/admin/cancel/final-approved')
assert cancel_response.status_code == 302, cancel_response.status_code

updated_requests = json.loads(Path('borrow_requests.json').read_text(encoding='utf-8'))
matching = next((req for req in updated_requests if req.get('id') == 'final-approved'), None)
assert matching is not None, 'Request missing after cancellation'
assert matching['status'] == 'cancelled', matching

save_json(REQUESTS_PATH, [])
save_json(BORROWED_PATH, [])
print('Verification OK: public page hides borrowed list, admin sees it, return date is captured, and approved requests can be cancelled.')
