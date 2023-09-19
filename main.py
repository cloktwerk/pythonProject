import pandas as pd
import matplotlib.pyplot as plt

# Загрузка данных из файла Excel
df = pd.read_excel('data.xlsx')

# Задача 1: Вычислите общую выручку за июль 2021 по сделкам, приход денежных средств которых не просрочен.
df['receiving_date'] = pd.to_datetime(df['receiving_date'], format='%Y-%m-%d', errors='coerce')
july_2021_revenue = df[(df['receiving_date'].dt.month == 7) &
                        (df['receiving_date'] <= df['receiving_date'].max()) &
                        (df['status'] == 'ОПЛАЧЕНО')]['sum'].sum()
print("Общая выручка за июль 2021 (без просроченных сделок):", july_2021_revenue)

# Задача 2: Как изменялась выручка компании за рассматриваемый период? Проиллюстрируйте графиком.

monthly_revenue = df.groupby(df['receiving_date'].dt.to_period("M"))['sum'].sum()
dates = monthly_revenue.index.to_timestamp()
plt.plot(dates, monthly_revenue.values, marker='o')
plt.plot(monthly_revenue.index, monthly_revenue.values, marker='o')
plt.xlabel('Дата')
plt.ylabel('Выручка')
plt.title('Изменение выручки компании по месяцам')
plt.xticks(rotation=45)
plt.grid(True)
plt.show()

# Задача 3: Кто из менеджеров привлек для компании больше всего денежных средств в сентябре 2021?
september_2021_revenue = df[(df['receiving_date'].dt.month == 9) &
                            (df['receiving_date'].dt.year == 2021)]['sum'].sum()
manager_with_highest_revenue = df[(df['receiving_date'].dt.month == 9) &
                                  (df['receiving_date'].dt.year == 2021)].groupby('sale')['sum'].sum().idxmax()
print(f"Менеджер, привлекший больше всего денежных средств в сентябре 2021: {manager_with_highest_revenue}")

# Задача 4: Какой тип сделок (новая/текущая) был преобладающим в октябре 2021?
october_2021_deals = df[(df['receiving_date'].dt.month == 10) &
                        (df['receiving_date'].dt.year == 2021)]
deal_type_counts = october_2021_deals['new/current'].value_counts()
dominant_deal_type = deal_type_counts.idxmax()
print(f"Преобладающий тип сделок в октябре 2021: {dominant_deal_type}")

# Задача 5: Сколько оригиналов договора по майским сделкам было получено в июне 2021?
may_2021_originals_received_in_june = df[(df['receiving_date'].dt.month == 5) &
                                          (df['receiving_date'].dt.year == 2021) &
                                          (df['receiving_date'].dt.month == 6)]['document'].sum()
print(f"Количество оригиналов договора по майским сделкам, полученных в июне 2021: {may_2021_originals_received_in_june}")

# Задача 6: Вычислите остаток каждого из менеджеров на 01.07.2021.
managers = df['sale'].unique()
manager_balances = {}
for manager in managers:
    manager_balance = df[(df['receiving_date'] < '2021-07-01') &
                         (df['sale'] == manager)]['sum'].sum()
    manager_balances[manager] = manager_balance

print("Остатки менеджеров на 01.07.2021:")
for manager, balance in manager_balances.items():
    print(f"{manager}: {balance}")


