# IMPORTS
import openpyxl
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import statsmodels.api as sm
from sklearn.linear_model import LinearRegression
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

# Data loading and preparation
df = pd.read_excel('krakow_real_estate_data.xlsx', header=1)

df['Transaction Date'] = pd.to_datetime(df['Transaction Date'])
df['Date_num'] = (df['Transaction Date'] - df['Transaction Date'].min()).dt.days

# Statistics
print("STATISTICS")
print(df['Unit Price [PLN/m2]'].describe())

# Analysis 1: Time trend
model_trend = LinearRegression()
model_trend.fit(df[['Date_num']], df['Unit Price [PLN/m2]'])

daily_trend = model_trend.coef_[0]
print(f"\nTREND: Daily price increase: {daily_trend:.2f} PLN/m2")
print(f"Estimated annual increase: {daily_trend * 365:.2f} PLN/m2")

plt.figure(figsize=(8,5))
sns.regplot(
    x='Date_num',
    y='Unit Price [PLN/m2]',
    data=df,
    line_kws={'color': 'red'}
)
plt.title('Price Growth Trend Over Time', fontweight='bold')
plt.xlabel('Days from the start of analysis')
plt.ylabel('Unit price [PLN/m2]')
plt.grid(True, linestyle='--', alpha=0.7)
plt.show()

# Analysis 2: Regression (Total Price vs Area)
X_area = df[['Usable Area [m2]']]
y_total = df['Total Price [PLN]']

model_area = LinearRegression()
model_area.fit(X_area, y_total)
r2_area = model_area.score(X_area, y_total)

print(f"\nREGRESSION: Total Price vs Area")
print(f"Each m2 increases the property price by: {model_area.coef_[0]:.2f} PLN")
print(f"Model fit (R2): {r2_area:.2f}")

plt.figure(figsize=(8,5))
sns.regplot(
    x='Usable Area [m2]',
    y='Total Price [PLN]',
    data=df,
    scatter_kws={'s': 40},
    line_kws={'color': 'red'}
)
plt.title('Regression: Total Price vs Area', fontweight='bold')
plt.xlabel('Area [m2]')
plt.ylabel('Total Price [PLN]')
plt.grid(True, linestyle='--', alpha=0.7)
plt.show()

# Analysis 3: K-means (Segmentation)
X_clus = df[['Unit Price [PLN/m2]', 'Location Rating']].copy()
scaler = StandardScaler()
X_scaled = scaler.fit_transform(X_clus)

kmeans = KMeans(n_clusters=3, random_state=42, n_init=10)
df['Segment_ID'] = kmeans.fit_predict(X_scaled)

segment_means = df.groupby('Segment_ID')['Unit Price [PLN/m2]'].mean().sort_values()

name_map = {
    segment_means.index[0]: 'Economy',
    segment_means.index[1]: 'Medium',
    segment_means.index[2]: 'Premium'
}

df['Segment'] = df['Segment_ID'].map(name_map)

print("\nCLASSIFICATION (Average prices in groups):")
print(df.groupby('Segment')['Unit Price [PLN/m2]'].mean().sort_values())

segment_order = ['Premium', 'Medium', 'Economy']
df['Segment'] = pd.Categorical(df['Segment'], categories=segment_order, ordered=True)

plt.figure(figsize=(8,5))
sns.scatterplot(
    x='Usable Area [m2]',
    y='Unit Price [PLN/m2]',
    hue='Segment',
    style='Location Rating',
    s=100,
    data=df,
    hue_order=segment_order
)

plt.title('Real Estate Market Segmentation (K-Means)', fontweight='bold')
plt.xlabel('Area [m2]')
plt.ylabel('Unit price [PLN/m2]')
plt.grid(True, linestyle='--', alpha=0.7)

plt.legend(bbox_to_anchor=(1.05, 1), loc='upper left', borderaxespad=0.)
plt.tight_layout()
plt.show()

# Analysis 4: Multiple Regression (Feature Weights)
X_multi = df[['Location Rating', 'Finishing Standard', 'Transport Accessibility']]
X_multi = sm.add_constant(X_multi)

y_multi = df['Unit Price [PLN/m2]']

model_weights = sm.OLS(y_multi, X_multi).fit()
print("\nFEATURE WEIGHTS (Multiple regression results):")
print(model_weights.params)

# Analysis 5: Feature Correlation (Correlation Matrix)
cols_corr = ['Location Rating', 'Finishing Standard',
             'Transport Accessibility', 'Usable Area [m2]',
             'Date_num', 'Unit Price [PLN/m2]']

cols_rename = {
    'Location Rating': 'Location',
    'Finishing Standard': 'Standard',
    'Transport Accessibility': 'Access',
    'Usable Area [m2]': 'Area',
    'Date_num': 'Time',
    'Unit Price [PLN/m2]': 'Price/m2'
}

corr_matrix = df[cols_corr].rename(columns=cols_rename).corr()

plt.figure(figsize=(10,6))
sns.heatmap(
    corr_matrix,
    annot=True,
    cmap='coolwarm',
    vmin=-1,
    vmax=1,
    center=0,
    fmt=".2f"
)
plt.title('Correlation Matrix of Price Factors', fontweight='bold')
plt.show()

# Correlation results
print("CORRELATION COEFFICIENTS WITH PRICE/M2")
print(corr_matrix['Price/m2'].sort_values(ascending=False))