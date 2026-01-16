from flask import Flask, render_template, request, redirect, url_for, flash, session
from flask_sqlalchemy import SQLAlchemy
from datetime import datetime
from functools import wraps
import os

app = Flask(__name__)
app.config['SECRET_KEY'] = 'aai-ims-secret-key-2026'
app.config['SQLALCHEMY_DATABASE_URI'] = os.environ.get('DATABASE_URL', 'sqlite:///assets.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

db = SQLAlchemy(app)

# ===================== MODELS =====================

class User(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    username = db.Column(db.String(80), unique=True, nullable=False)
    password = db.Column(db.String(120), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # 'admin' or 'employee'
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)


class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.String(50), unique=True, nullable=False)
    name = db.Column(db.String(120), nullable=False)
    email = db.Column(db.String(120))
    mobile = db.Column(db.String(20))
    department = db.Column(db.String(80), nullable=False)
    designation = db.Column(db.String(80))
    
    assets = db.relationship('Asset', backref='employee_obj', lazy=True)
    assignments = db.relationship('AssetAssignment', backref='employee', lazy=True)
    user = db.relationship('User', backref='employee', uselist=False)


class Asset(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    asset_name = db.Column(db.String(120), nullable=False)
    serial_number = db.Column(db.String(120), nullable=False, unique=True)
    asset_tag = db.Column(db.String(120))
    category = db.Column(db.String(80), nullable=False)
    location = db.Column(db.String(120), nullable=False)
    site = db.Column(db.String(120))
    status = db.Column(db.String(50), nullable=False, default='In Store')
    vendor_name = db.Column(db.String(120))
    purchase_cost = db.Column(db.Float, default=0.0)
    warranty_expiry = db.Column(db.String(20))
    assigned_employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=True)
    assigned_to = db.Column(db.String(120))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    assignments = db.relationship('AssetAssignment', backref='asset', lazy=True)
    reports = db.relationship('AssetReport', backref='asset', lazy=True)


class AssetAssignment(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    asset_id = db.Column(db.Integer, db.ForeignKey('asset.id'), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    assigned_date = db.Column(db.DateTime, default=datetime.utcnow)
    returned_date = db.Column(db.DateTime, nullable=True)
    notes = db.Column(db.Text, default='Issued to employee')


class AssetReport(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    asset_id = db.Column(db.Integer, db.ForeignKey('asset.id'), nullable=False)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    report_type = db.Column(db.String(50))  # 'issue', 'maintenance', 'comment'
    message = db.Column(db.Text, nullable=False)
    status = db.Column(db.String(20), default='pending')  # 'pending', 'resolved'
    created_at = db.Column(db.DateTime, default=datetime.utcnow)
    
    reporter = db.relationship('Employee', backref='reports')


# ===================== AUTH DECORATORS =====================

def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            flash('Please login first.', 'danger')
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated_function


def admin_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if 'user_id' not in session:
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('Admin access required.', 'danger')
            return redirect(url_for('employee_dashboard'))
        return f(*args, **kwargs)
    return decorated_function


# ===================== HELPERS =====================

def compute_stats():
    total = Asset.query.count()
    in_store = Asset.query.filter_by(status='In Store').count()
    deployed = Asset.query.filter_by(status='Deployed').count()
    maintenance = Asset.query.filter_by(status='In Maintenance').count()
    
    in_store_val = db.session.query(db.func.sum(Asset.purchase_cost))\
        .filter(Asset.status == 'In Store').scalar() or 0
    deployed_val = db.session.query(db.func.sum(Asset.purchase_cost))\
        .filter(Asset.status == 'Deployed').scalar() or 0
    
    return {
        "total_assets": total,
        "in_store": in_store,
        "deployed": deployed,
        "in_maintenance": maintenance,
        "in_store_value": in_store_val,
        "deployed_value": deployed_val
    }


def get_employee_by_name(name):
    if not name:
        return None
    return Employee.query.filter(Employee.name == name).first()


# ===================== AUTH ROUTES =====================

@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        
        user = User.query.filter_by(username=username, password=password).first()
        
        if user:
            session['user_id'] = user.id
            session['username'] = user.username
            session['role'] = user.role
            session['employee_id'] = user.employee_id
            
            flash(f'Welcome {username}!', 'success')
            
            if user.role == 'admin':
                return redirect(url_for('index'))
            else:
                return redirect(url_for('employee_dashboard'))
        else:
            flash('Invalid credentials!', 'danger')
    
    return render_template('login.html')

@app.route('/reports')
@login_required
@admin_required
def admin_reports():
    # Get all reports with asset and employee info
    reports = (
        db.session.query(AssetReport, Asset, Employee)
        .join(Asset, Asset.id == AssetReport.asset_id)
        .join(Employee, Employee.id == AssetReport.employee_id)
        .order_by(AssetReport.created_at.desc())
        .all()
    )
    
    # Filter by status if requested
    status_filter = request.args.get('status', '')
    if status_filter:
        reports = [r for r in reports if r[0].status == status_filter]
    
    return render_template('admin_reports.html', reports=reports, status_filter=status_filter)


@app.route('/resolve-report/<int:report_id>', methods=['POST'])
@login_required
@admin_required
def resolve_report(report_id):
    report = AssetReport.query.get_or_404(report_id)
    report.status = 'resolved'
    db.session.commit()
    flash('Report marked as resolved!', 'success')
    return redirect(url_for('admin_reports'))

@app.route('/logout')
def logout():
    session.clear()
    flash('Logged out successfully.', 'success')
    return redirect(url_for('login'))


# ===================== ADMIN ROUTES =====================

@app.route('/')
@login_required
@admin_required
def index():
    stats = compute_stats()
    recent_assets = Asset.query.order_by(Asset.created_at.desc()).limit(10).all()
    return render_template('index.html', stats=stats, recent_assets=recent_assets)


@app.route('/assets')
@login_required
@admin_required
def list_assets():
    search = request.args.get('search', '').strip()
    category = request.args.get('category', '')
    status = request.args.get('status', '')
    
    query = Asset.query
    
    if search:
        like = f"%{search}%"
        query = query.filter(
            db.or_(Asset.asset_name.ilike(like),
                   Asset.serial_number.ilike(like))
        )
    
    if category:
        query = query.filter_by(category=category)
    
    if status:
        query = query.filter_by(status=status)
    
    assets = query.order_by(Asset.created_at.desc()).all()
    return render_template(
        'assets.html',
        assets=assets,
        search_query=search,
        category=category,
        status=status
    )


@app.route('/add-asset', methods=['GET', 'POST'])
@login_required
@admin_required
def add_asset():
    employees = Employee.query.order_by(Employee.name).all()
    
    if request.method == 'POST':
        assigned_name = request.form.get('assigned_to') or None
        emp = get_employee_by_name(assigned_name)
        assigned_id = emp.id if emp else None
        
        asset = Asset(
            asset_name=request.form['asset_name'],
            serial_number=request.form['serial_number'],
            asset_tag=request.form.get('asset_tag') or None,
            category=request.form['category'],
            location=request.form['location'],
            site=request.form.get('site') or None,
            status=request.form['status'],
            vendor_name=request.form.get('vendor_name') or None,
            purchase_cost=float(request.form.get('purchase_cost') or 0),
            warranty_expiry=request.form.get('warranty_expiry') or None,
            assigned_employee_id=assigned_id,
            assigned_to=assigned_name
        )
        
        db.session.add(asset)
        db.session.commit()
        
        if assigned_id:
            assignment = AssetAssignment(
                asset_id=asset.id,
                employee_id=assigned_id,
                notes=f"Asset assigned to {assigned_name}"
            )
            db.session.add(assignment)
            db.session.commit()
        
        flash('Asset added successfully!', 'success')
        return redirect(url_for('list_assets'))
    
    return render_template('add_asset.html', employees=employees)


@app.route('/edit-asset/<int:asset_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    employees = Employee.query.order_by(Employee.name).all()
    
    if request.method == 'POST':
        assigned_name = request.form.get('assigned_to') or None
        emp = get_employee_by_name(assigned_name)
        assigned_id = emp.id if emp else None
        
        asset.asset_name = request.form['asset_name']
        asset.serial_number = request.form['serial_number']
        asset.asset_tag = request.form.get('asset_tag') or None
        asset.category = request.form['category']
        asset.location = request.form['location']
        asset.site = request.form.get('site') or None
        asset.status = request.form['status']
        asset.vendor_name = request.form.get('vendor_name') or None
        asset.purchase_cost = float(request.form.get('purchase_cost') or 0)
        asset.warranty_expiry = request.form.get('warranty_expiry') or None
        asset.assigned_employee_id = assigned_id
        asset.assigned_to = assigned_name
        
        db.session.commit()
        
        if assigned_id:
            assignment = AssetAssignment(
                asset_id=asset.id,
                employee_id=assigned_id,
                notes="Asset re-assigned"
            )
            db.session.add(assignment)
            db.session.commit()
        
        flash('Asset updated!', 'success')
        return redirect(url_for('list_assets'))
    
    return render_template('edit_asset.html', asset=asset, employees=employees)


@app.route('/delete-asset/<int:asset_id>', methods=['POST'])
@login_required
@admin_required
def delete_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    db.session.delete(asset)
    db.session.commit()
    flash('Asset deleted.', 'danger')
    return redirect(url_for('list_assets'))


@app.route('/employees')
@login_required
@admin_required
def list_employees():
    employees = Employee.query.order_by(Employee.name).all()
    
    enriched = []
    for e in employees:
        enriched.append({
            "id": e.id,
            "employee_id": e.employee_id,
            "name": e.name,
            "email": e.email or '-',
            "mobile": e.mobile or '-',
            "department": e.department,
            "designation": e.designation or '-',
            "assets_assigned": len(e.assets)
        })
    
    return render_template('employees.html', employees=enriched)


@app.route('/add-employee', methods=['GET', 'POST'])
@login_required
@admin_required
def add_employee():
    if request.method == 'POST':
        emp = Employee(
            employee_id=request.form['employee_id'],
            name=request.form['name'],
            email=request.form.get('email') or None,
            mobile=request.form.get('mobile') or None,
            department=request.form['department'],
            designation=request.form.get('designation') or None
        )
        db.session.add(emp)
        db.session.commit()
        
        # Create employee user account
        username = request.form.get('username') or emp.employee_id
        password = request.form.get('password', 'password123')
        
        user = User(
            username=username,
            password=password,
            role='employee',
            employee_id=emp.id
        )
        db.session.add(user)
        db.session.commit()
        
        flash(f'Employee added! Login: {username} / {password}', 'success')
        return redirect(url_for('list_employees'))
    
    return render_template('add_employee.html')


@app.route('/edit-employee/<int:emp_id>', methods=['GET', 'POST'])
@login_required
@admin_required
def edit_employee(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    
    if request.method == 'POST':
        emp.employee_id = request.form['employee_id']
        emp.name = request.form['name']
        emp.email = request.form.get('email') or None
        emp.mobile = request.form.get('mobile') or None
        emp.department = request.form['department']
        emp.designation = request.form.get('designation') or None
        
        db.session.commit()
        flash('Employee updated!', 'success')
        return redirect(url_for('list_employees'))
    
    return render_template('edit_employee.html', employee=emp)


@app.route('/delete-employee/<int:emp_id>', methods=['POST'])
@login_required
@admin_required
def delete_employee(emp_id):
    emp = Employee.query.get_or_404(emp_id)
    db.session.delete(emp)
    db.session.commit()
    flash('Employee deleted.', 'danger')
    return redirect(url_for('list_employees'))


@app.route('/tracking')
@login_required
@admin_required
def tracking():
    assignments = (
        db.session.query(AssetAssignment, Asset, Employee)
        .join(Asset, Asset.id == AssetAssignment.asset_id)
        .join(Employee, Employee.id == AssetAssignment.employee_id)
        .order_by(AssetAssignment.assigned_date.desc())
        .all()
    )
    return render_template('tracking.html', assignments=assignments)


# ===================== EMPLOYEE ROUTES =====================

@app.route('/employee-dashboard')
@login_required
def employee_dashboard():
    if session.get('role') != 'employee':
        return redirect(url_for('index'))
    
    emp_id = session.get('employee_id')
    employee = Employee.query.get_or_404(emp_id)
    
    my_assets = Asset.query.filter_by(assigned_employee_id=emp_id).all()
    my_reports = AssetReport.query.filter_by(employee_id=emp_id)\
        .order_by(AssetReport.created_at.desc()).limit(10).all()
    
    return render_template('employee_dashboard.html', 
                         employee=employee, 
                         assets=my_assets,
                         reports=my_reports)


@app.route('/report-asset/<int:asset_id>', methods=['GET', 'POST'])
@login_required
def report_asset(asset_id):
    asset = Asset.query.get_or_404(asset_id)
    emp_id = session.get('employee_id')
    
    if request.method == 'POST':
        report = AssetReport(
            asset_id=asset_id,
            employee_id=emp_id,
            report_type=request.form['report_type'],
            message=request.form['message']
        )
        db.session.add(report)
        db.session.commit()
        
        flash('Report submitted successfully!', 'success')
        return redirect(url_for('employee_dashboard'))
    
    return render_template('report_asset.html', asset=asset)


# ===================== MAIN =====================

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
        
        # Create admin if not exists
        admin = User.query.filter_by(username='admin').first()
        if not admin:
            admin = User(username='admin', password='admin123', role='admin')
            db.session.add(admin)
            db.session.commit()
            print("✅ Admin created: admin / admin123")
    
    app.run(debug=True)
