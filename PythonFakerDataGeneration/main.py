'''
Author: Wade Zhong wzhong@hso.com
Date: 2024-04-24 15:01:12
LastEditTime: 2024-04-24 15:40:20
LastEditors: Wade Zhong wzhong@hso.com
Description: 
FilePath: \JavaScriptWPSMacros\PythonFakerDataGeneration\main.py
Copyright (c) 2024 by Wade Zhong wzhong@hso.com, All Rights Reserved. 
'''
import pandas as pd
from faker import Faker
import random
import uuid

# 初始化 Faker 库，设置为中国区域
fake = Faker('zh_CN')


# 生成产品系列数据，只有5个
productSeriesLength = 5
product_series = pd.DataFrame({
    'series_id': [str(uuid.uuid4()) for _ in range(productSeriesLength)],
    'series_name': [fake.catch_phrase() for _ in range(productSeriesLength)]
})

# 生成产品数据
productsLength = 100
products = pd.DataFrame({
    'product_id': [str(uuid.uuid4()) for _ in range(productsLength)],
    'series_id': [random.choice(product_series['series_id']) for _ in range(productsLength)],  # 每个产品都有一个产品系列
    'product_name': [fake.catch_phrase() for _ in range(productsLength)],
    'price': [random.uniform(10.0, 200.0) for _ in range(productsLength)]
})

# 生成用户数据
usersLength = 100
# 定义一个邮箱域名列表
email_domains = ['@microsoft.com', '@google.com', '@qq.com', '@163.com']
users = pd.DataFrame({
    'user_id': [str(uuid.uuid4()) for _ in range(usersLength)],  # 使用UUID作为用户ID
    'province': [fake.province() for _ in range(usersLength)],  # 省份
    'city': [fake.city() for _ in range(usersLength)],  # 城市
    'district': [fake.district() for _ in range(usersLength)],  # 区县
    'address': [fake.address() for _ in range(usersLength)],  # 区县
    'name': [fake.name() for _ in range(usersLength)],  # 中文名字
    'phone_number': [fake.phone_number() for _ in range(usersLength)],  # 电话号码
    'credit_card_number': [fake.credit_card_number() for _ in range(usersLength)],  # 银行卡号
    'company': [fake.company() for _ in range(usersLength)],  # 公司
    'job': [fake.job() for _ in range(usersLength)],  # 工作
    'gender': [random.choice(["男","女"]) for _ in range(usersLength)],  # 性别
    'IDNumber': [fake.ssn(min_age=18, max_age=90) for _ in range(usersLength)],  # 性别
    'email': [fake.user_name() + random.choice(email_domains) for _ in range(usersLength)]
})

# 生成订单数据
ordersLength = 100
orders = pd.DataFrame({
    'order_id': [str(uuid.uuid4()) for _ in range(ordersLength)],
    'product_id': [random.choice(products['product_id']) for _ in range(ordersLength)],
    'user_id': [random.choice(users['user_id']) for _ in range(ordersLength)],
    'quantity': [random.randint(1, 10) for _ in range(ordersLength)],
    'date_time': [fake.date_time() for _ in range(ordersLength)]
})

# 生成订单详情数据
orderDetailsLength = 800
order_details = pd.DataFrame({
    'order_id': [random.choice(orders['order_id']) for _ in range(orderDetailsLength)],
    'product_id': [random.choice(products['product_id']) for _ in range(orderDetailsLength)],
    'quantity': [random.randint(1, 10) for _ in range(orderDetailsLength)],
    'price': [random.uniform(10.0, 200.0) for _ in range(orderDetailsLength)]
})
# 按照订单ID排序订单详情数据
order_details = order_details.sort_values(by='order_id')

# 生成文章数据
articlesLength = 100
articles = pd.DataFrame({
    'article_id': [str(uuid.uuid4()) for _ in range(articlesLength)],
    'title': [fake.sentence() for _ in range(articlesLength)],
    'content': [fake.text() for _ in range(articlesLength)]
})

# 创建一个Excel写入器
writer = pd.ExcelWriter('data.xlsx')

# 将数据写入Excel文件的不同sheet
product_series.to_excel(writer, sheet_name='ProductSeries', index=False)
products.to_excel(writer, sheet_name='Products', index=False)
users.to_excel(writer, sheet_name='Users', index=False)
orders.to_excel(writer, sheet_name='Orders', index=False)
order_details.to_excel(writer, sheet_name='OrderDetails', index=False)
articles.to_excel(writer, sheet_name='Articles', index=False)

# 保存Excel文件
writer.close()
