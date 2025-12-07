from datetime import datetime, date, timedelta
import re

def validate_courier_data(form_data):
    """Валидация данных курьерской доставки"""
    errors = []
    
    # Валидация веса
    try:
        weight = float(form_data.get('weight', 0))
        if weight <= 0:
            errors.append('❌ Вес должен быть больше 0 кг')
        elif weight > 1000:
            errors.append('❌ Вес не может превышать 1000 кг')
    except ValueError:
        errors.append('❌ Неверный формат веса')
    
    # Валидация габаритов
    dimensions = ['length', 'width', 'height']
    for dim in dimensions:
        try:
            value = float(form_data.get(dim, 0))
            if value <= 0:
                errors.append(f'❌ {dim.capitalize()} должен быть больше 0 см')
            elif value > 500:
                errors.append(f'❌ {dim.capitalize()} не может превышать 500 см')
        except ValueError:
            errors.append(f'❌ Неверный формат {dim}')
    
    # Валидация дат
    try:
        dispatch_date = datetime.strptime(form_data.get('dispatch_date', ''), '%Y-%m-%d').date()
        delivery_date = datetime.strptime(form_data.get('delivery_date', ''), '%Y-%m-%d').date()
        today = date.today()
        
        if dispatch_date < today - timedelta(days=1):
            errors.append('❌ Дата отправления не может быть в прошлом')
        if delivery_date < dispatch_date:
            errors.append('❌ Дата получения не может быть раньше даты отправления')
        if delivery_date > today + timedelta(days=365):
            errors.append('❌ Дата получения не может быть больше чем через год')
    except ValueError:
        errors.append('❌ Неверный формат даты')
    
    # Валидация стоимости доставки
    try:
        delivery_cost = float(form_data.get('delivery_cost', 0))
        if delivery_cost < 0:
            errors.append('❌ Стоимость доставки не может быть отрицательной')
        elif delivery_cost > 1000000:
            errors.append('❌ Стоимость доставки слишком высокая')
    except ValueError:
        errors.append('❌ Неверный формат стоимости доставки')
    
    # Валидация паспортных данных отправителя
    errors.extend(validate_passport_data(form_data, 'sender'))
    
    # Валидация паспортных данных получателя
    errors.extend(validate_passport_data(form_data, 'receiver'))
    
    # Валидация телефона курьера
    phone = form_data.get('courier_phone', '').strip()
    if not validate_phone(phone):
        errors.append('❌ Неверный формат телефона курьера')
    
    # Валидация ФИО
    for field in ['sender_name', 'receiver_name', 'courier_name']:
        name = form_data.get(field, '').strip()
        if len(name) < 2:
            errors.append(f'❌ Поле "{field}" должно содержать минимум 2 символа')
        elif not re.match(r'^[А-Яа-яЁёA-Za-z\s\-\.]+$', name):
            errors.append(f'❌ Неверный формат ФИО в поле "{field}"')
    
    # Валидация адресов
    for field in ['sender_address', 'receiver_address']:
        address = form_data.get(field, '').strip()
        if len(address) < 5:
            errors.append(f'❌ Адрес "{field}" слишком короткий')
    
    return {
        'valid': len(errors) == 0,
        'errors': errors
    }

def validate_passport_data(form_data, prefix):
    """Валидация паспортных данных"""
    errors = []
    
    # Серия паспорта (4 цифры)
    series = form_data.get(f'{prefix}_passport_series', '').strip()
    if not re.match(r'^\d{4}$', series):
        errors.append(f'❌ Серия паспорта {prefix} должна содержать 4 цифры')
    
    # Номер паспорта (6 цифр)
    number = form_data.get(f'{prefix}_passport_number', '').strip()
    if not re.match(r'^\d{6}$', number):
        errors.append(f'❌ Номер паспорта {prefix} должен содержать 6 цифр')
    
    # Дата рождения
    try:
        birth_date = datetime.strptime(form_data.get(f'{prefix}_birth_date', ''), '%Y-%m-%d').date()
        today = date.today()
        age = today.year - birth_date.year - ((today.month, today.day) < (birth_date.month, birth_date.day))
        
        if age < 14:
            errors.append(f'❌ Возраст {prefix} должен быть не менее 14 лет')
        elif age > 120:
            errors.append(f'❌ Некорректная дата рождения {prefix}')
        
        # Проверка что дата рождения не в будущем
        if birth_date > today:
            errors.append(f'❌ Дата рождения {prefix} не может быть в будущем')
    except ValueError:
        errors.append(f'❌ Неверный формат даты рождения {prefix}')
    
    # Пол
    gender = form_data.get(f'{prefix}_gender', '').strip()
    if gender not in ['М', 'Ж']:
        errors.append(f'❌ Неверное значение пола {prefix}. Допустимо: М или Ж')
    
    return errors

def validate_course_data(form_data):
    """Валидация данных курсов повышения квалификации"""
    errors = []
    
    # Валидация названия курса
    course_name = form_data.get('course_name', '').strip()
    if len(course_name) < 5:
        errors.append('❌ Название курса должно содержать минимум 5 символов')
    
    # Валидация количества часов
    try:
        hours = int(form_data.get('hours', 0))
        if hours <= 0:
            errors.append('❌ Количество часов должно быть больше 0')
        elif hours > 1000:
            errors.append('❌ Количество часов не может превышать 1000')
    except ValueError:
        errors.append('❌ Неверный формат количества часов')
    
    # Валидация дат
    try:
        start_date = datetime.strptime(form_data.get('start_date', ''), '%Y-%m-%d').date()
        end_date = datetime.strptime(form_data.get('end_date', ''), '%Y-%m-%d').date()
        today = date.today()
        
        if start_date < today:
            errors.append('❌ Дата начала не может быть в прошлом')
        if end_date < start_date:
            errors.append('❌ Дата окончания не может быть раньше даты начала')
        if (end_date - start_date).days > 365:
            errors.append('❌ Длительность курса не может превышать 1 год')
        
        # Проверка дедлайна регистрации
        reg_deadline = form_data.get('registration_deadline')
        if reg_deadline:
            reg_date = datetime.strptime(reg_deadline, '%Y-%m-%d').date()
            if reg_date > start_date:
                errors.append('❌ Дедлайн регистрации не может быть позже даты начала курса')
    except ValueError:
        errors.append('❌ Неверный формат даты')
    
    # Валидация преподавателя
    teacher_name = form_data.get('teacher_name', '').strip()
    if len(teacher_name) < 2:
        errors.append('❌ ФИО преподавателя должно содержать минимум 2 символа')
    
    # Валидация отдела преподавателя
    teacher_department = form_data.get('teacher_department', '').strip()
    if len(teacher_department) < 2:
        errors.append('❌ Название отдела должно содержать минимум 2 символа')
    
    # Валидация цены
    try:
        price = float(form_data.get('price', 0))
        if price < 0:
            errors.append('❌ Цена не может быть отрицательной')
        elif price > 1000000:
            errors.append('❌ Цена слишком высокая')
    except ValueError:
        errors.append('❌ Неверный формат цены')
    
    # Валидация максимального количества участников
    try:
        max_participants = int(form_data.get('max_participants', 30))
        if max_participants <= 0:
            errors.append('❌ Максимальное количество участников должно быть больше 0')
        elif max_participants > 1000:
            errors.append('❌ Максимальное количество участников не может превышать 1000')
    except ValueError:
        errors.append('❌ Неверный формат максимального количества участников')
    
    # Валидация email преподавателя
    teacher_email = form_data.get('teacher_email', '').strip()
    if teacher_email and not validate_email(teacher_email):
        errors.append('❌ Неверный формат email преподавателя')
    
    # Валидация телефона преподавателя
    teacher_phone = form_data.get('teacher_phone', '').strip()
    if teacher_phone and not validate_phone(teacher_phone):
        errors.append('❌ Неверный формат телефона преподавателя')
    
    # Валидация сотрудников
    employee_count = 0
    for i in range(1, 4):
        emp_name = form_data.get(f'employee_{i}_name', '').strip()
        emp_position = form_data.get(f'employee_{i}_position', '').strip()
        
        if emp_name and not emp_position:
            errors.append(f'❌ Для сотрудника {i} указано ФИО, но не указана должность')
        elif not emp_name and emp_position:
            errors.append(f'❌ Для сотрудника {i} указана должность, но не указано ФИО')
        elif emp_name and emp_position:
            employee_count += 1
            if len(emp_name) < 2:
                errors.append(f'❌ ФИО сотрудника {i} должно содержать минимум 2 символа')
            if len(emp_position) < 2:
                errors.append(f'❌ Должность сотрудника {i} должна содержать минимум 2 символа')
            
            # Валидация email сотрудника
            emp_email = form_data.get(f'employee_{i}_email', '').strip()
            if emp_email and not validate_email(emp_email):
                errors.append(f'❌ Неверный формат email сотрудника {i}')
    
    if employee_count == 0:
        errors.append('❌ Необходимо указать хотя бы одного сотрудника')
    
    return {
        'valid': len(errors) == 0,
        'errors': errors
    }

def validate_phone(phone):
    """Валидация номера телефона"""
    # Российские номера: +7XXXXXXXXXX или 8XXXXXXXXXX
    pattern = r'^(\+7|8)\d{10}$'
    return bool(re.match(pattern, phone))

def validate_email(email):
    """Валидация email"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))