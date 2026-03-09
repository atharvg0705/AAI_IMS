from app import app, db, User, Employee

with app.app_context():
    existing = Employee.query.filter_by(employee_id='ADMIN001').first()
    if existing:
        u = User.query.filter_by(role='admin').first()
        u.employee_id = existing.id
        db.session.commit()
        print(f"Admin already linked to Employee ID: {existing.id}")
    else:
        admin_emp = Employee(
            employee_id='ADMIN001',
            name='Admin',
            department='Administration',
            email='admin@aai.aero'
        )
        db.session.add(admin_emp)
        db.session.flush()
        u = User.query.filter_by(role='admin').first()
        u.employee_id = admin_emp.id
        db.session.commit()
        print(f"Done! Admin linked to Employee ID: {admin_emp.id}")
