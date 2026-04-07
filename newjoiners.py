import gspread
from google.oauth2.service_account import Credentials
import pandas as pd
from prophet import Prophet

# ----------------------
# Google Sheets Auth
# ----------------------
scopes = ["https://www.googleapis.com/auth/spreadsheets"]
creds = Credentials.from_service_account_file("credentials.json", scopes=scopes)
client = gspread.authorize(creds)

# Open workbook
sheet_id = "1akuaXxI5j4lzQTPj8F8fFNXQxThFjUeupSbNazCn25k"
Workbook = client.open_by_key(sheet_id)

# Read employee data
sheet = Workbook.worksheet("Employees ")
data = sheet.get_all_records()
df = pd.DataFrame(data)
print(df.tail(5))
print("Row count:", len(df))
from datetime import datetime
print("Last run:", datetime.now())


# ----------------------
# Convert Join_Date and group by month
# ----------------------
df['Join_Date'] = pd.to_datetime(df['Join_Date'], dayfirst=True, errors='coerce')

df['YearMonth'] = df['Join_Date'].dt.to_period('M')
monthly_joiners = df.groupby('YearMonth').size().reset_index(name='y')

# Prophet expects ds, y
monthly_joiners['ds'] = monthly_joiners['YearMonth'].dt.to_timestamp()
monthly_joiners = monthly_joiners[['ds', 'y']]

# Fill missing months with 0
monthly_joiners = monthly_joiners.set_index('ds').asfreq('MS').fillna(0).reset_index()

# ----------------------
# Forecast with Prophet
# ----------------------
model = Prophet()
model.fit(monthly_joiners)

# Make future dataframe (6 months ahead)
future = model.make_future_dataframe(periods=6, freq='MS')  # Month start
forecast = model.predict(future)

# ----------------------
# Sort and clean values
# ----------------------
forecast = forecast.sort_values('ds').reset_index(drop=True)

# Clip negative values
for col in ['yhat', 'yhat_lower', 'yhat_upper']:
    forecast[col] = forecast[col].apply(lambda x: max(x, 0))

# Round to integers
forecast['yhat'] = forecast['yhat'].round(0).astype(int)
forecast['yhat_lower'] = forecast['yhat_lower'].round(0).astype(int)
forecast['yhat_upper'] = forecast['yhat_upper'].round(0).astype(int)

# Pick last 6 months chronologically
forecast_output = forecast[['ds', 'yhat', 'yhat_lower', 'yhat_upper']].tail(6)
forecast_output = forecast_output.sort_values('ds').reset_index(drop=True)

# ----------------------
# Add numeric month for Looker Studio sorting
# ----------------------
forecast_output['Month_Num'] = forecast_output['ds'].dt.month
forecast_output['Month_Name'] = forecast_output['ds'].dt.strftime('%b')  # Apr, May, etc.

# ----------------------
# Push to Google Sheets
# ----------------------
try:
    forecast_sheet = Workbook.worksheet("Forecast")
except:
    forecast_sheet = Workbook.add_worksheet(title="Forecast", rows="100", cols="10")

forecast_sheet.clear()

# Convert ds to string for gspread (ISO format for Looker Studio)
forecast_output['ds'] = forecast_output['ds'].dt.strftime('%Y-%m-%d')

forecast_sheet.update(
    [forecast_output.columns.values.tolist()] + forecast_output.values.tolist()
)

print("✅ Forecast updated successfully")
print(forecast_output)