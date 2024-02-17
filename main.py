import pandas as pd
import pygsheets

# Load the data
central_orders = pd.read_csv('datasets/Orders_Central.csv', encoding='latin1')
orders_2015 = pd.read_csv('datasets/orders_south_2015.csv', encoding='latin1')
orders_2016 = pd.read_csv('datasets/Orders_Central.csv', encoding='latin1')
orders_2017 = pd.read_csv('datasets/Orders_Central.csv', encoding='latin1')
orders_2018 = pd.read_csv('datasets/Orders_Central.csv', encoding='latin1')
return_reasons_new = pd.read_excel('datasets/return reasons_new.xlsx')

# Function to format dates
def date_format(row):
    order_date = pd.to_datetime(f"{row['Order Year']}-{row['Order Month']}-{row['Order Day']}").date()
    ship_date = pd.to_datetime(f"{row['Ship Year']}-{row['Ship Month']}-{row['Ship Day']}").date()
    return pd.Series({'Order Date': order_date, 'Ship Date': ship_date})

# Data processing for 2015
orders_2015[['Order Date', 'Ship Date']] = orders_2015[['Order Date', 'Ship Date']].apply(pd.to_datetime, errors='coerce')
orders_2015.rename(columns={'Product Name': 'Product','Discount':'Discounts'}, inplace=True)
orders_2015 = orders_2015.drop(columns=['Region'])
orders_2015['Discounts'] = orders_2015['Discounts'].fillna(0)
orders_2015.reset_index(drop=True, inplace=True)

# Data processing for 2016
orders_2016[['Order Date', 'Ship Date']] = orders_2016.apply(lambda row: date_format(row), axis=1)
orders_2016[['Order Date', 'Ship Date']] = orders_2016[['Order Date', 'Ship Date']].apply(pd.to_datetime, errors='coerce')
orders_2016 = orders_2016[pd.DatetimeIndex(orders_2016['Order Date']).year == 2016]
orders_2016 = orders_2016.drop(columns=['Order Year', 'Order Month', 'Order Day', 'Ship Year', 'Ship Month', 'Ship Day'])
orders_2016['Discounts'] = orders_2016['Discounts'].fillna(0)
orders_2016.reset_index(drop=True, inplace=True)

# data_2018_from_orders_2016=orders_2016[pd.DatetimeIndex(orders_2016['Order Date']).year == 2018]
# data_2017_from_orders_2016=orders_2016[pd.DatetimeIndex(orders_2016['Order Date']).year == 2017]
# data_2015_from_orders_2016=orders_2016[pd.DatetimeIndex(orders_2016['Order Date']).year == 2015]

# Data processing for 2017
orders_2017[['Order Date', 'Ship Date']] = orders_2017.apply(lambda row: date_format(row), axis=1)
orders_2017[['Order Date', 'Ship Date']] = orders_2017[['Order Date', 'Ship Date']].apply(pd.to_datetime, errors='coerce')
orders_2017 = orders_2017[pd.DatetimeIndex(orders_2017['Order Date']).year == 2017]
orders_2017 = orders_2017.drop(columns=['Order Year', 'Order Month', 'Order Day', 'Ship Year', 'Ship Month', 'Ship Day'])
orders_2017['Discounts'] = orders_2017['Discounts'].fillna(0)
orders_2017.reset_index(drop=True, inplace=True)

# data_2018_from_orders_2017=orders_2017[pd.DatetimeIndex(orders_2017['Order Date']).year == 2018]
# data_2016_from_orders_2017=orders_2017[pd.DatetimeIndex(orders_2017['Order Date']).year == 2016]
# data_2015_from_orders_2017=orders_2017[pd.DatetimeIndex(orders_2017['Order Date']).year == 2015]

# Data processing for 2018
orders_2018[['Order Date', 'Ship Date']] = orders_2018.apply(lambda row: date_format(row), axis=1)
orders_2018[['Order Date', 'Ship Date']] = orders_2018[['Order Date', 'Ship Date']].apply(pd.to_datetime, errors='coerce')
orders_2018 = orders_2018[pd.DatetimeIndex(orders_2018['Order Date']).year == 2018]
orders_2018 = orders_2018.drop(columns=['Order Year', 'Order Month', 'Order Day', 'Ship Year', 'Ship Month', 'Ship Day'])
orders_2018['Discounts'] = orders_2018['Discounts'].fillna(0)
orders_2018.reset_index(drop=True, inplace=True)

# data_2017_from_orders_2018=orders_2018[pd.DatetimeIndex(orders_2018['Order Date']).year == 2017]
# data_2016_from_orders_2018=orders_2018[pd.DatetimeIndex(orders_2018['Order Date']).year == 2016]
# data_2015_from_orders_2018=orders_2018[pd.DatetimeIndex(orders_2018['Order Date']).year == 2015]

#Data processing for Central Orders
central_orders[['Order Date', 'Ship Date']] = central_orders.apply(lambda row: date_format(row), axis=1)
central_orders[['Order Date', 'Ship Date']] = central_orders[['Order Date', 'Ship Date']].apply(pd.to_datetime, errors='coerce')
central_orders = central_orders[(central_orders['Order Date'].dt.year >= 2015) & (central_orders['Order Date'].dt.year <= 2018)]
central_orders = central_orders.drop(columns=['Order Year', 'Order Month', 'Order Day', 'Ship Year', 'Ship Month', 'Ship Day'])
null_columns_central_orders = central_orders.columns[central_orders.isnull().any()]
central_orders['Discounts'] = central_orders['Discounts'].fillna(0)
central_orders.reset_index(drop=True, inplace=True)

# Check for invalid ship dates where the order year is greater than the ship year
dfs = {
   'central_orders': central_orders,
    'orders_2015': orders_2015,
    'orders_2016': orders_2016,
    'orders_2017': orders_2017,
    'orders_2018': orders_2018
}
for name, df in dfs.items():
    invalid_ship_dates = df[df['Order Date'].dt.year> df['Ship Date'].dt.year]
    if not invalid_ship_dates.empty:
        print(f"In {name}, there are rows where Order Year is greater than Ship Year:")
        print(invalid_ship_dates)
    else:
        print(f"In {name}, there are no rows where Order Year is greater than Ship Year.")


#are_equal = data_2018_from_orders_2017.equals(data_2018_from_orders_2016)
#matching_orders = orders_2018[orders_2018['Order ID'].isin(data_2018_from_orders_2016['Order ID'])]

# Merging and cleaning the data

orders = pd.concat([orders_2015, orders_2016, orders_2017, orders_2018], ignore_index=True)
cleaned_central_orders = central_orders[~central_orders['Order ID'].isin(orders['Order ID'])]
merged_orders = pd.concat([orders, cleaned_central_orders], ignore_index=True)


#Data processing for Return reasons new dataset
return_reasons_new['Order Date'] = pd.to_datetime(return_reasons_new['Order Date'])
return_reasons_new = return_reasons_new[(return_reasons_new['Order Date'].dt.year >= 2015) & (return_reasons_new['Order Date'].dt.year <= 2018)]
return_reasons_new.loc[:, 'Order Year'] = return_reasons_new['Order Date'].dt.year

return_reasons_new = pd.merge(return_reasons_new, merged_orders[['Category','Sub-Category']], on='Sub-Category', how='left')
return_reasons_new=return_reasons_new.drop_duplicates()
return_reasons_new=return_reasons_new.reset_index()
return_reasons_new.to_excel('output files/return reasons.xlsx')


# Return rate analysis

shipped_orders_per_year = merged_orders.groupby(merged_orders['Order Date'].dt.year)['Order ID'].count()
returned_orders_per_year = return_reasons_new.groupby(return_reasons_new['Order Date'].dt.year)['Order ID'].count()
returned_vs_shipped = pd.DataFrame({'Shipped Orders': shipped_orders_per_year, 'Returned Orders': returned_orders_per_year})
returned_vs_shipped['Return Rate (%)'] = round((returned_vs_shipped['Returned Orders'] / returned_vs_shipped['Shipped Orders']) * 100,2)
returned_vs_shipped = returned_vs_shipped.fillna(0)

returned_vs_shipped.to_excel('output files/frequency of order returns.xlsx')

# Product category analysis

most_returned_products = return_reasons_new['Sub-Category'].value_counts().reset_index()
most_returned_products.columns = ['Sub-Category', 'Count']

most_returned_products.to_excel('output files/most return products.xlsx')

merged = pd.merge(return_reasons_new, merged_orders[['Product ID', 'Category']], on='Product ID', how='left')

