from flask import Flask, render_template, request, redirect, url_for, flash, send_file
from pymongo import MongoClient
from datetime import datetime, date, timedelta
import os
from bson.objectid import ObjectId
import reports
from validation import validate_courier_data, validate_course_data
from export import export_to_pdf, export_to_docx
import re
import io
import random
import string

app = Flask(__name__)
app.secret_key = 'your-secret-key-here-change-in-production'

# Подключение к MongoDB
client = MongoClient('mongodb://localhost:27017/')
db = client['documents_db']

# Коллекции
courier_collection = db['courier_deliveries']
courses_collection = db['qualification_courses']

# Контекстный процессор для передачи данных во все шаблоны
@app.context_processor
def inject_today():
    return {'today': datetime.now().strftime('%Y-%m-%d')}

# Главная страница
@app.route('/')
def index():
    return render_template('index.html')

# ========== КУРЬЕРСКАЯ ДОСТАВКА ==========

# Список всех посылок
@app.route('/courier')
def courier_list():
    parcels = list(courier_collection.find().sort('created_at', -1))
    return render_template('courier_list.html', parcels=parcels)

# Добавление новой посылки
@app.route('/courier/add', methods=['GET', 'POST'])
def add_courier():
    if request.method == 'POST':
        # Валидация данных
        validation_result = validate_courier_data(request.form)
        
        if not validation_result['valid']:
            for error in validation_result['errors']:
                flash(error, 'danger')
            return render_template('courier_form.html', 
                                 parcel=request.form, 
                                 action='Добавить',
                                 errors=validation_result['errors'])
        
        # Подготовка данных
        parcel = {
            'sender': {
                'full_name': request.form['sender_name'].strip(),
                'address': request.form['sender_address'].strip(),
                'passport': {
                    'series': request.form['sender_passport_series'].strip(),
                    'number': request.form['sender_passport_number'].strip(),
                    'birth_date': request.form['sender_birth_date'],
                    'gender': request.form['sender_gender']
                }
            },
            'receiver': {
                'full_name': request.form['receiver_name'].strip(),
                'address': request.form['receiver_address'].strip(),
                'passport': {
                    'series': request.form['receiver_passport_series'].strip(),
                    'number': request.form['receiver_passport_number'].strip(),
                    'birth_date': request.form['receiver_birth_date'],
                    'gender': request.form['receiver_gender']
                }
            },
            'parcel': {
                'weight': float(request.form['weight']),
                'dimensions': {
                    'length': float(request.form['length']),
                    'width': float(request.form['width']),
                    'height': float(request.form['height'])
                },
                'description': request.form.get('description', '').strip(),
                'fragile': request.form.get('fragile') == 'on',
                'insured': request.form.get('insured') == 'on'
            },
            'courier': {
                'name': request.form['courier_name'].strip(),
                'phone': request.form['courier_phone'].strip(),
                'vehicle': request.form.get('courier_vehicle', '').strip(),
                'company': request.form.get('courier_company', '').strip()
            },
            'dates': {
                'dispatch_date': request.form['dispatch_date'],
                'delivery_date': request.form['delivery_date'],
                'actual_delivery_date': None
            },
            'status': request.form['status'],
            'created_at': datetime.now(),
            'tracking_number': generate_tracking_number(),
            'delivery_cost': float(request.form.get('delivery_cost', 0))
        }
        
        courier_collection.insert_one(parcel)
        flash('Посылка успешно добавлена! Трек номер: ' + parcel['tracking_number'], 'success')
        return redirect(url_for('courier_list'))
    
    return render_template('courier_form.html', parcel=None, action='Добавить')

# Редактирование посылки
@app.route('/courier/edit/<id>', methods=['GET', 'POST'])
def edit_courier(id):
    parcel = courier_collection.find_one({'_id': ObjectId(id)})
    
    if request.method == 'POST':
        # Валидация данных
        validation_result = validate_courier_data(request.form)
        
        if not validation_result['valid']:
            for error in validation_result['errors']:
                flash(error, 'danger')
            return render_template('courier_form.html', 
                                 parcel={**request.form, '_id': id}, 
                                 action='Редактировать',
                                 errors=validation_result['errors'])
        
        update_data = {
            'sender': {
                'full_name': request.form['sender_name'].strip(),
                'address': request.form['sender_address'].strip(),
                'passport': {
                    'series': request.form['sender_passport_series'].strip(),
                    'number': request.form['sender_passport_number'].strip(),
                    'birth_date': request.form['sender_birth_date'],
                    'gender': request.form['sender_gender']
                }
            },
            'receiver': {
                'full_name': request.form['receiver_name'].strip(),
                'address': request.form['receiver_address'].strip(),
                'passport': {
                    'series': request.form['receiver_passport_series'].strip(),
                    'number': request.form['receiver_passport_number'].strip(),
                    'birth_date': request.form['receiver_birth_date'],
                    'gender': request.form['receiver_gender']
                }
            },
            'parcel': {
                'weight': float(request.form['weight']),
                'dimensions': {
                    'length': float(request.form['length']),
                    'width': float(request.form['width']),
                    'height': float(request.form['height'])
                },
                'description': request.form.get('description', '').strip(),
                'fragile': request.form.get('fragile') == 'on',
                'insured': request.form.get('insured') == 'on'
            },
            'courier': {
                'name': request.form['courier_name'].strip(),
                'phone': request.form['courier_phone'].strip(),
                'vehicle': request.form.get('courier_vehicle', '').strip(),
                'company': request.form.get('courier_company', '').strip()
            },
            'dates': {
                'dispatch_date': request.form['dispatch_date'],
                'delivery_date': request.form['delivery_date'],
                'actual_delivery_date': parcel.get('dates', {}).get('actual_delivery_date')
            },
            'status': request.form['status'],
            'updated_at': datetime.now(),
            'delivery_cost': float(request.form.get('delivery_cost', 0))
        }
        
        # Если статус изменился на "Доставлено", устанавливаем фактическую дату доставки
        if request.form['status'] == 'Доставлено' and parcel.get('status') != 'Доставлено':
            update_data['dates']['actual_delivery_date'] = datetime.now().strftime('%Y-%m-%d')
        
        courier_collection.update_one({'_id': ObjectId(id)}, {'$set': update_data})
        flash('Посылка успешно обновлена!', 'success')
        return redirect(url_for('courier_list'))
    
    # Преобразуем данные для отображения в форме
    if parcel:
        form_data = {
            '_id': str(parcel['_id']),
            'sender_name': parcel['sender']['full_name'],
            'sender_address': parcel['sender']['address'],
            'sender_passport_series': parcel['sender']['passport']['series'],
            'sender_passport_number': parcel['sender']['passport']['number'],
            'sender_birth_date': parcel['sender']['passport']['birth_date'],
            'sender_gender': parcel['sender']['passport']['gender'],
            'receiver_name': parcel['receiver']['full_name'],
            'receiver_address': parcel['receiver']['address'],
            'receiver_passport_series': parcel['receiver']['passport']['series'],
            'receiver_passport_number': parcel['receiver']['passport']['number'],
            'receiver_birth_date': parcel['receiver']['passport']['birth_date'],
            'receiver_gender': parcel['receiver']['passport']['gender'],
            'weight': parcel['parcel']['weight'],
            'length': parcel['parcel']['dimensions']['length'],
            'width': parcel['parcel']['dimensions']['width'],
            'height': parcel['parcel']['dimensions']['height'],
            'description': parcel['parcel'].get('description', ''),
            'fragile': parcel['parcel'].get('fragile', False),
            'insured': parcel['parcel'].get('insured', False),
            'courier_name': parcel['courier']['name'],
            'courier_phone': parcel['courier']['phone'],
            'courier_vehicle': parcel['courier'].get('vehicle', ''),
            'courier_company': parcel['courier'].get('company', ''),
            'dispatch_date': parcel['dates']['dispatch_date'],
            'delivery_date': parcel['dates']['delivery_date'],
            'delivery_cost': parcel.get('delivery_cost', 0),
            'status': parcel['status']
        }
        return render_template('courier_form.html', parcel=form_data, action='Редактировать')
    
    return render_template('courier_form.html', parcel=parcel, action='Редактировать')

# Удаление посылки
@app.route('/courier/delete/<id>')
def delete_courier(id):
    courier_collection.delete_one({'_id': ObjectId(id)})
    flash('Посылка успешно удалена!', 'success')
    return redirect(url_for('courier_list'))

# Просмотр деталей посылки
@app.route('/courier/view/<id>')
def view_courier(id):
    parcel = courier_collection.find_one({'_id': ObjectId(id)})
    if not parcel:
        flash('Посылка не найдена!', 'danger')
        return redirect(url_for('courier_list'))
    return render_template('courier_view.html', parcel=parcel)

# ========== ПОВЫШЕНИЕ КВАЛИФИКАЦИИ ==========

# Список всех курсов
@app.route('/courses')
def courses_list():
    courses = list(courses_collection.find().sort('created_at', -1))
    return render_template('courses_list.html', courses=courses)

# Добавление нового курса
@app.route('/courses/add', methods=['GET', 'POST'])
def add_course():
    if request.method == 'POST':
        # Валидация данных
        validation_result = validate_course_data(request.form)
        
        if not validation_result['valid']:
            for error in validation_result['errors']:
                flash(error, 'danger')
            return render_template('courses_form.html', 
                                 course=request.form, 
                                 action='Добавить',
                                 errors=validation_result['errors'])
        
        # Получаем до 3 сотрудников
        employees = []
        for i in range(1, 4):
            emp_name = request.form.get(f'employee_{i}_name', '').strip()
            emp_position = request.form.get(f'employee_{i}_position', '').strip()
            emp_department = request.form.get(f'employee_{i}_department', '').strip()
            emp_email = request.form.get(f'employee_{i}_email', '').strip()
            if emp_name and emp_position:
                employees.append({
                    'name': emp_name,
                    'position': emp_position,
                    'department': emp_department,
                    'email': emp_email
                })
        
        course = {
            'course_name': request.form['course_name'].strip(),
            'course_code': generate_course_code(),
            'teacher': {
                'name': request.form['teacher_name'].strip(),
                'department': request.form['teacher_department'].strip(),
                'qualification': request.form.get('teacher_qualification', '').strip(),
                'email': request.form.get('teacher_email', '').strip(),
                'phone': request.form.get('teacher_phone', '').strip()
            },
            'dates': {
                'start_date': request.form['start_date'],
                'end_date': request.form['end_date'],
                'registration_deadline': request.form.get('registration_deadline', '')
            },
            'hours': int(request.form['hours']),
            'price': float(request.form.get('price', 0)),
            'location': request.form.get('location', '').strip(),
            'max_participants': int(request.form.get('max_participants', 30)),
            'current_participants': 0,
            'employees': employees,
            'status': request.form['status'],
            'description': request.form.get('description', '').strip(),
            'created_at': datetime.now(),
            'category': request.form.get('category', 'Общий').strip()
        }
        
        courses_collection.insert_one(course)
        flash('Курс успешно добавлен! Код курса: ' + course['course_code'], 'success')
        return redirect(url_for('courses_list'))
    
    return render_template('courses_form.html', course=None, action='Добавить')

# Редактирование курса
@app.route('/courses/edit/<id>', methods=['GET', 'POST'])
def edit_course(id):
    course = courses_collection.find_one({'_id': ObjectId(id)})
    
    if request.method == 'POST':
        # Валидация данных
        validation_result = validate_course_data(request.form)
        
        if not validation_result['valid']:
            for error in validation_result['errors']:
                flash(error, 'danger')
            return render_template('courses_form.html', 
                                 course={**request.form, '_id': id}, 
                                 action='Редактировать',
                                 errors=validation_result['errors'])
        
        employees = []
        for i in range(1, 4):
            emp_name = request.form.get(f'employee_{i}_name', '').strip()
            emp_position = request.form.get(f'employee_{i}_position', '').strip()
            emp_department = request.form.get(f'employee_{i}_department', '').strip()
            emp_email = request.form.get(f'employee_{i}_email', '').strip()
            if emp_name and emp_position:
                employees.append({
                    'name': emp_name,
                    'position': emp_position,
                    'department': emp_department,
                    'email': emp_email
                })
        
        update_data = {
            'course_name': request.form['course_name'].strip(),
            'teacher': {
                'name': request.form['teacher_name'].strip(),
                'department': request.form['teacher_department'].strip(),
                'qualification': request.form.get('teacher_qualification', '').strip(),
                'email': request.form.get('teacher_email', '').strip(),
                'phone': request.form.get('teacher_phone', '').strip()
            },
            'dates': {
                'start_date': request.form['start_date'],
                'end_date': request.form['end_date'],
                'registration_deadline': request.form.get('registration_deadline', '')
            },
            'hours': int(request.form['hours']),
            'price': float(request.form.get('price', 0)),
            'location': request.form.get('location', '').strip(),
            'max_participants': int(request.form.get('max_participants', 30)),
            'current_participants': course.get('current_participants', 0),
            'employees': employees,
            'status': request.form['status'],
            'description': request.form.get('description', '').strip(),
            'updated_at': datetime.now(),
            'category': request.form.get('category', 'Общий').strip()
        }
        
        courses_collection.update_one({'_id': ObjectId(id)}, {'$set': update_data})
        flash('Курс успешно обновлен!', 'success')
        return redirect(url_for('courses_list'))
    
    # Преобразуем данные для отображения в форме
    if course:
        form_data = {
            '_id': str(course['_id']),
            'course_name': course['course_name'],
            'teacher_name': course['teacher']['name'],
            'teacher_department': course['teacher']['department'],
            'teacher_qualification': course['teacher'].get('qualification', ''),
            'teacher_email': course['teacher'].get('email', ''),
            'teacher_phone': course['teacher'].get('phone', ''),
            'start_date': course['dates']['start_date'],
            'end_date': course['dates']['end_date'],
            'registration_deadline': course['dates'].get('registration_deadline', ''),
            'hours': course['hours'],
            'price': course.get('price', 0),
            'location': course.get('location', ''),
            'max_participants': course.get('max_participants', 30),
            'status': course['status'],
            'description': course.get('description', ''),
            'category': course.get('category', 'Общий')
        }
        
        # Добавляем данные сотрудников
        for i, emp in enumerate(course.get('employees', []), 1):
            form_data[f'employee_{i}_name'] = emp.get('name', '')
            form_data[f'employee_{i}_position'] = emp.get('position', '')
            form_data[f'employee_{i}_department'] = emp.get('department', '')
            form_data[f'employee_{i}_email'] = emp.get('email', '')
        
        return render_template('courses_form.html', course=form_data, action='Редактировать')
    
    return render_template('courses_form.html', course=course, action='Редактировать')

# Удаление курса
@app.route('/courses/delete/<id>')
def delete_course(id):
    courses_collection.delete_one({'_id': ObjectId(id)})
    flash('Курс успешно удален!', 'success')
    return redirect(url_for('courses_list'))

# Просмотр деталей курса
@app.route('/courses/view/<id>')
def view_course(id):
    course = courses_collection.find_one({'_id': ObjectId(id)})
    if not course:
        flash('Курс не найден!', 'danger')
        return redirect(url_for('courses_list'))
    return render_template('course_view.html', course=course)

# ========== ОТЧЕТЫ И ЭКСПОРТ ==========

@app.route('/reports')
def show_reports():
    reports_data = reports.generate_reports(courier_collection, courses_collection)
    return render_template('reports.html', reports=reports_data)

@app.route('/export/pdf/<report_type>/<report_name>')
def export_pdf(report_type, report_name):
    reports_data = reports.generate_reports(courier_collection, courses_collection)
    
    if report_type not in ['courier', 'courses']:
        flash('Неверный тип отчета', 'danger')
        return redirect(url_for('show_reports'))
    
    pdf_data = export_to_pdf(reports_data, report_type, report_name)
    
    if pdf_data:
        filename = f'report_{report_type}_{report_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.pdf'
        return send_file(
            io.BytesIO(pdf_data),
            download_name=filename,
            mimetype='application/pdf',
            as_attachment=True
        )
    else:
        flash('Ошибка при создании PDF', 'danger')
        return redirect(url_for('show_reports'))

@app.route('/export/docx/<report_type>/<report_name>')
def export_docx(report_type, report_name):
    reports_data = reports.generate_reports(courier_collection, courses_collection)
    
    if report_type not in ['courier', 'courses']:
        flash('Неверный тип отчета', 'danger')
        return redirect(url_for('show_reports'))
    
    docx_data = export_to_docx(reports_data, report_type, report_name)
    
    if docx_data:
        filename = f'report_{report_type}_{report_name}_{datetime.now().strftime("%Y%m%d_%H%M%S")}.docx'
        return send_file(
            io.BytesIO(docx_data),
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True
        )
    else:
        flash('Ошибка при создании DOCX', 'danger')
        return redirect(url_for('show_reports'))

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========

def generate_tracking_number():
    """Генерация уникального трек-номера"""
    timestamp = datetime.now().strftime('%Y%m%d')
    random_part = ''.join(random.choices(string.ascii_uppercase + string.digits, k=8))
    return f"TRK{timestamp}{random_part}"

def generate_course_code():
    """Генерация уникального кода курса"""
    timestamp = datetime.now().strftime('%y%m')
    random_part = ''.join(random.choices(string.ascii_uppercase, k=3))
    return f"COURSE{timestamp}{random_part}"

# Кастомные фильтры для Jinja2
@app.template_filter('sum_employees')
def sum_employees_filter(courses):
    """Суммирует количество сотрудников во всех курсах"""
    total = 0
    for course in courses:
        total += len(course.get('employees', []))
    return total

# ========== ОШИБКИ ==========

@app.errorhandler(404)
def page_not_found(e):
    return render_template('404.html'), 404

@app.errorhandler(500)
def internal_server_error(e):
    return render_template('500.html'), 500

if __name__ == '__main__':
    # Создаем индексы для ускорения поиска
    courier_collection.create_index([('tracking_number', 1)], unique=True)
    courier_collection.create_index([('status', 1)])
    courier_collection.create_index([('dates.dispatch_date', -1)])
    
    courses_collection.create_index([('course_code', 1)], unique=True)
    courses_collection.create_index([('status', 1)])
    courses_collection.create_index([('dates.start_date', -1)])
    
    app.run(debug=True, port=5000)