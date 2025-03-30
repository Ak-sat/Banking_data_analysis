import pandas as pd

#Provide a meaningful treatment to all values where age is less than 18
key = pd.read_excel(r'Credit Banking_Project1 (1).xls')
# Filter the dataset to select rows where age is less than 18
filtered_key = key[key['Age'] < 18]
filtered_key.loc[:, 'Age'] = key.Age.mean()
# Update the original dataset with the treated values
key.update(filtered_key)
# Print the updated dataset
#print(key.to_string())

# Calculate monthly spend of each customer
key1 = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Spend')
key1['Date'] = pd.to_datetime(key1['Month'], format='%d-%m-%Y')
# Extract the month from the 'Date' column
key1['Month'] = key1['Date'].dt.to_period('M')
# Group the data by customer and month, and calculate the total spend for each customer in each month
monthly_spend = key1.groupby(['Customer', 'Month'])['Amount'].sum()
# Print the monthly spend of each customer
#print(monthly_spend.to_string())

# Calculate monthly repayment of each customer
key2 = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Repayment')
key2['Date'] = pd.to_datetime(key2['Month'], format='%d-%m-%Y')
# Extract the month from the 'Date' column
key2['Month'] = key2['Date'].dt.to_period('M')
# Group the data by customer and month, and calculate the total repayment for each customer in each month
monthly_repayment = key2.groupby(['Customer', 'Month'])['Amount'].sum()
# Print the monthly repayment of each customer
#print(monthly_repayment.to_string())

#Highest paying 10 customers
key3 = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Repayment')
print("Highest Paying 10 Customers:")
sorted_key3 = key3.sort_values('Amount', ascending=False)
# Get the top 10 highest paying customers
top_10_customers = sorted_key3.head(10)
# Print the highest paying customers
#print(top_10_customers)

#ppl in which segment are spending more money
segment_data = pd.read_excel(r'Credit Banking_Project1 (1).xls',
                             'Customer Acqusition')
spend_data = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Spend')
# Merge the age group and spending data based on a common column
merged_data = pd.merge(segment_data, spend_data, on='Customer')
# Calculate the total spend for each segment
segment_spending = merged_data.groupby('Segment')['Amount'].sum()
# Find the segment with the highest spend
most_spending_segment = segment_spending.idxmax()
# Calculate the total amount spent across all segments
total_spending = segment_spending.sum()
# Print the results
#print("Segment spending:")
#print(segment_spending)
#print("\nThe segment with the highest spend is:", most_spending_segment)

#ppl of which age group are spending more money
age_data = pd.read_excel(r'Credit Banking_Project1 (1).xls',
                         'Customer Acqusition')
spending_data = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Spend')
# Merge the age group and spending data based on a common column
merged_data = pd.merge(age_data, spending_data, on='Customer')
# Group by age group and calculate total spending
agegroup_spending = merged_data.groupby('Age')['Amount'].sum()
# Find the age group with the highest spending
max_spending_agegroup = agegroup_spending.idxmax()
# Get the total spending for the max spending age group
max_agegroup_spending = agegroup_spending.loc[max_spending_agegroup]
#print("Customers of age group", max_spending_agegroup, "are spending more money.")
#print("Total spending for this age group:", max_agegroup_spending)

#Which is the most profitable segment?
#segment_data = pd.read_excel(r'Credit Banking_Project1 (3) .10.xls',
#                            'Customer Acqusition')
# Group the data by segment and calculate the total profit for each segment
#seg_profit = segment_data.groupby('Segment')['Limit'].sum()
# Find the segment with the highest profit
#most_profitable_segment = seg_profit.idxmax()
#print("Most profitable segment:", most_profitable_segment)

#Category in which people are spending more money
# Read the spend data with category and amount month-wise
spend_data = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Spend')
# Calculate the total spend for each category
category_spending = spend_data.groupby('Type')['Amount'].sum()
# Find the category with the highest spend
most_spending_category = category_spending.idxmax()
# Calculate the total amount spent across all categories
total_spending = category_spending.sum()
# Print the results
#print("Category spending:")
#print(category_spending)
#print("\nThe category with the highest spend is:", most_spending_category)
#print("Total amount spent across all categories:", total_spending)

#Identity where the repayment is more than the spend then give them a credit of 2% of their surplus amount in the next month billing.
key7 = pd.read_excel(r'Edulyt excel.xlsx', 'Sheet2')
# Calculate surplus amount for customers with repayment more than spend
# Calculate the surplus amount (repayment - spend)
key7['Surplus'] = key7['Repayment Amount'] - key7['Spend Amount']
# Identify customers with surplus amount and give them a credit of 2% in the next month billing
key7.loc[key7['Surplus'] > 0, 'Credit'] = key7['Surplus'] * 0.02
# Print the updated dataset with credit information
#print(key7.to_string())

#Monthly profit for the bank.
# Read the repayment and spend data from Excel sheets
repayment_data = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Repayment')
spend_data = pd.read_excel(r'Credit Banking_Project1 (1).xls', 'Spend')
# Group the repayment data by month and calculate the total repayment amount
repayment_monthly = repayment_data.groupby('Month')['Amount'].sum()
# Group the spend data by month and calculate the total spend amount
spend_monthly = spend_data.groupby('Month')['Amount'].sum()
# Group the data by month and calculate the total profit for each month
# Calculate the monthly profit for the bank
# Calculate the monthly profit for the bank
monthly_profit = spend_monthly.subtract(repayment_monthly, fill_value=0)
# Group the data by month and calculate the total profit for each month
monthly_profit = monthly_profit.sort_values(ascending=False)
#monthly_profit = spend_monthly - repayment_monthly
# Print the monthly profit
#print(monthly_profit.to_string())

key8 = pd.read_excel(r'Edulyt excel.xlsx', 'Sheet2')
# Calculate monthly profit for the bank
monthly_profit = key8['Repayment Amount'].sum() - key8['Spend Amount'].sum()
#print(key8.to_string())

#Impose an interest rate of 2.9% for each customer for any due amount.
key9 = pd.read_excel(r'Edulyt excel.xlsx', 'Sheet2')
# Calculate interest amount and update repayment amount
interest_rate = 0.029  # 2.9%
# Calculate interest amount for each row
key9['Interest'] = key9.apply(lambda row: max(
  row['Spend Amount'] - row['Repayment Amount'], 0) * interest_rate,
                              axis=1)
# Update repayment amount by adding the interest amount
key9['Repayment Amount'] = key9['Repayment Amount'] + key9['Interest']
# Save the updated dataset to a new Excel file
key9.to_excel('updated_dataset.xlsx', index=False)
#print(key9.to_string())
