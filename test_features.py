import os

os.environ.setdefault('LAB_ADMIN_USERNAME', 'admin')
os.environ.setdefault('LAB_ADMIN_PASSWORD', 'TestPassword123!')

from app import app, get_audit_events, get_overdue_requests, get_zero_stock_alerts

assert callable(get_zero_stock_alerts)
assert callable(get_overdue_requests)
assert callable(get_audit_events)

with app.test_client() as client:
    response = client.get('/admin')
    assert response.status_code in (200, 302)

    admin_login = client.post(
        '/login',
        data={'username': 'admin', 'password': 'TestPassword123!'},
    )
    assert admin_login.status_code in (200, 302)

    history_page = client.get('/admin/history')
    assert history_page.status_code in (200, 302)

    reminder_response = client.post('/admin/remind')
    assert reminder_response.status_code in (200, 302)

print('Feature checks OK')
