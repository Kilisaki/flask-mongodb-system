from pymongo import MongoClient
from datetime import datetime, timedelta

def generate_reports(courier_collection, courses_collection):
    """Генерация отчетов"""
    
    reports_data = {}
    
    # 1. Отчет по курьерской доставке
    reports_data['courier_reports'] = {
        # Все посылки с весом более 5 кг
        'heavy_parcels': list(courier_collection.find({
            'parcel.weight': {'$gt': 5}
        }).sort('parcel.weight', -1).limit(100)),
        
        # Посылки в процессе доставки
        'in_transit': list(courier_collection.find({
            'status': {'$in': ['В пути', 'Обработка', 'В пункте выдачи']}
        }).sort('dates.delivery_date', 1).limit(100)),
        
        # Посылки за последние 7 дней
        'last_week': list(courier_collection.find({
            'dates.dispatch_date': {
                '$gte': (datetime.now() - timedelta(days=7)).strftime('%Y-%m-%d')
            }
        }).sort('dates.dispatch_date', -1).limit(100)),
        
        # Посылки от определенного отправителя (пример)
        'by_sender': list(courier_collection.find({
            'sender.full_name': {'$regex': 'Иванов', '$options': 'i'}
        }).limit(50)),
        
        # Статистика по курьерам
        'courier_stats': list(courier_collection.aggregate([
            {'$match': {'courier.name': {'$ne': None, '$ne': ''}}},
            {'$group': {
                '_id': '$courier.name',
                'count': {'$sum': 1},
                'total_weight': {'$sum': '$parcel.weight'},
                'total_cost': {'$sum': '$delivery_cost'}
            }},
            {'$sort': {'count': -1}},
            {'$limit': 20}
        ])),
        
        # Все посылки (ограниченное количество)
        'all': list(courier_collection.find().sort('created_at', -1).limit(50))
    }
    
    # 2. Отчет по курсам повышения квалификации
    reports_data['courses_reports'] = {
        # Предстоящие курсы
        'upcoming_courses': list(courses_collection.find({
            'dates.start_date': {'$gte': datetime.now().strftime('%Y-%m-%d')},
            'status': {'$in': ['Запланирован', 'Набор']}
        }).sort('dates.start_date', 1).limit(100)),
        
        # Курсы с количеством часов более 40
        'long_courses': list(courses_collection.find({
            'hours': {'$gt': 40}
        }).sort('hours', -1).limit(100)),
        
        # Курсы определенного преподавателя
        'by_teacher': list(courses_collection.find({
            'teacher.name': {'$regex': 'Петров', '$options': 'i'}
        }).limit(50)),
        
        # Курсы с заполненными группами (3 сотрудника)
        'full_courses': list(courses_collection.find({
            'employees.2': {'$exists': True}
        }).limit(100)),
        
        # Статистика по отделам
        'department_stats': list(courses_collection.aggregate([
            {'$unwind': '$employees'},
            {'$group': {
                '_id': '$employees.department',
                'employee_count': {'$sum': 1},
                'course_count': {'$addToSet': '$course_name'}
            }},
            {'$project': {
                'department': '$_id',
                'employee_count': 1,
                'course_count': {'$size': '$course_count'}
            }},
            {'$sort': {'employee_count': -1}},
            {'$limit': 20}
        ])),
        
        # Все курсы (ограниченное количество)
        'all': list(courses_collection.find().sort('created_at', -1).limit(50))
    }
    
    # 3. Общая статистика
    reports_data['general_stats'] = {
        'total_parcels': courier_collection.count_documents({}),
        'total_courses': courses_collection.count_documents({}),
        'parcels_in_transit': courier_collection.count_documents({
            'status': {'$in': ['В пути', 'Обработка', 'В пункте выдачи']}
        }),
        'upcoming_courses_count': courses_collection.count_documents({
            'dates.start_date': {'$gte': datetime.now().strftime('%Y-%m-%d')}
        }),
        'total_delivery_cost': list(courier_collection.aggregate([
            {'$group': {
                '_id': None,
                'total': {'$sum': '$delivery_cost'}
            }}
        ]))[0]['total'] if list(courier_collection.aggregate([{'$group': {'_id': None, 'total': {'$sum': '$delivery_cost'}}}])) else 0,
        'total_course_price': list(courses_collection.aggregate([
            {'$group': {
                '_id': None,
                'total': {'$sum': '$price'}
            }}
        ]))[0]['total'] if list(courses_collection.aggregate([{'$group': {'_id': None, 'total': {'$sum': '$price'}}}])) else 0
    }
    
    return reports_data

if __name__ == '__main__':
    # Тестирование модуля
    client = MongoClient('mongodb://localhost:27017/')
    db = client['documents_db']
    
    courier_collection = db['courier_deliveries']
    courses_collection = db['qualification_courses']
    
    reports_result = generate_reports(courier_collection, courses_collection)
    
    print("=== ТЕСТИРОВАНИЕ ОТЧЕТОВ ===")
    print(f"Всего посылок в системе: {reports_result['general_stats']['total_parcels']}")
    print(f"Всего курсов в системе: {reports_result['general_stats']['total_courses']}")
    print(f"Посылок в пути: {reports_result['general_stats']['parcels_in_transit']}")
    print(f"Предстоящих курсов: {reports_result['general_stats']['upcoming_courses_count']}")
    
    print("\n=== ОТЧЕТ ПО КУРЬЕРСКОЙ ДОСТАВКЕ ===")
    print(f"Тяжелые посылки (>5 кг): {len(reports_result['courier_reports']['heavy_parcels'])}")
    print(f"Посылок в пути: {len(reports_result['courier_reports']['in_transit'])}")
    print(f"Посылок за последнюю неделю: {len(reports_result['courier_reports']['last_week'])}")
    
    print("\n=== ОТЧЕТ ПО КУРСАМ ===")
    print(f"Предстоящих курсов: {len(reports_result['courses_reports']['upcoming_courses'])}")
    print(f"Длительных курсов (>40 часов): {len(reports_result['courses_reports']['long_courses'])}")
    print(f"Курсов с полными группами: {len(reports_result['courses_reports']['full_courses'])}")