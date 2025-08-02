import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import mplcyberpunk
plt.style.use('cyberpunk')

# Load the Excel file
df = pd.read_excel("restaurants_data.xlsx")

# Display the first 5 rows
print(df.head())


print()

# Descriptive statistics for delivery time
print("Delivery Time Statistics:")
print(df["Delivery Time"].describe())

# Descriptive statistics for rating
print("\n Rating Statistics:")
print(df["Rating"].describe())

# Count of restaurants with fast delivery (<= 45 minutes)
fast_delivery_count = (df["Delivery Time"] <= 45).sum()
print(f"\nNumber of restaurants with delivery time <= 45 minutes: {fast_delivery_count}")

# Count of restaurants with high rating (>= 4.5)
high_rating_count = (df["Rating"] >= 4.5).sum()
print(f"\nNumber of restaurants with rating >= 4.5: {high_rating_count}")

print("\nTop 10 Resrtaurants By ratings:")
top_10 = df.sort_values(by="Rating", ascending=False).head(10)
print(top_10[["Restaurant Name", "Rating", "Delivery Time", "Address"]])

print("\nWorest 10 Resrtaurants By ratings:")
worst_10 = df[df["Rating"] > 0].sort_values(by="Rating").head(10)
print(worst_10[["Restaurant Name", "Rating", "Delivery Time", "Address"]])

bins = [0, 30, 45, 60, 1000]
labels = ["Fast (<=30)", "Medium (31-45)", "Slow (46-60)", "Very Slow (>60)"]
df["Delivery Category"] = pd.cut(df["Delivery Time"], bins=bins, labels=labels)

fig, axes = plt.subplots(nrows=2, ncols=2, figsize=(15, 10))
plt.subplots_adjust(hspace=0.4, wspace=0.3)
sns.set_theme(style="whitegrid")

sns.countplot(data=df, x="Delivery Time", palette="Blues", ax=axes[0, 0])
axes[0, 0].set_title("Number of Restaurants by Delivery Time")
axes[0, 0].set_xlabel("Delivery Time")
axes[0, 0].set_ylabel("Number of Restaurants")
axes[0, 0].grid(axis='y', linestyle='--', alpha=0.7)

sns.countplot(data=df, x="Rating", palette="Blues", ax=axes[0, 1])
axes[0, 1].set_title("Number of Restaurants by Rating")
axes[0, 1].set_xlabel("Rating")
axes[0, 1].set_ylabel("Number of Restaurants")
axes[0, 1].grid(axis='y', linestyle='--', alpha=0.7)

sns.boxplot(data=df, x="Delivery Category", y="Rating", palette="Set2", ax=axes[1, 0])
axes[1, 0].set_title("Rating by Delivery Time Category")
axes[1, 0].set_xlabel("Delivery Time Category")
axes[1, 0].set_ylabel("Rating")
axes[1, 0].grid(axis='y', linestyle='--', alpha=0.7)

sns.scatterplot(data=df, x="Rates", y="Rating", alpha=0.6, ax=axes[1, 1])
axes[1, 1].set_title("Rating vs Number of Reviews")
axes[1, 1].set_xlabel("Number of Reviews")
axes[1, 1].set_ylabel("Rating")
axes[1, 1].grid(True)

plt.suptitle("Restaurant Data Visualization", fontsize=16 , color="White" )
plt.tight_layout(rect=[0, 0, 1, 0.96])
plt.show()

summary = {
    "Average Rating": [df["Rating"].mean()],
    "Average Delivery Time": [df["Delivery Time"].mean()],
    "Fast Deliveries (≤ 45 mins)": [(df["Delivery Time"] <= 45).sum()],
    "High Ratings (≥ 4.5)": [(df["Rating"] >= 4.5).sum()]
}


summary_df = pd.DataFrame(summary)
output_path = "summary_report.xlsx"
with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
    summary_df.to_excel(writer, index=False, sheet_name="Summary")
    workbook = writer.book
    worksheet = writer.sheets["Summary"]

    header_format = workbook.add_format({
        'bold': True,
        'bg_color': '#4472C4',
        'font_color': 'white',
        'align': 'center',
        'valign': 'vcenter',
        'border': 1
    })

    number_format = workbook.add_format({'num_format': '0.00', 'border': 1})
    int_format = workbook.add_format({'num_format': '0', 'border': 1})
    cell_format = workbook.add_format({'border': 1})

    for col_num, value in enumerate(summary_df.columns.values):
        worksheet.write(0, col_num, value, header_format)
        worksheet.set_column(col_num, col_num, max(20, len(value) + 5))

    for row_num in range(1, 2): 
        worksheet.write(row_num, 0, summary_df.at[0, "Average Rating"], number_format)
        worksheet.write(row_num, 1, summary_df.at[0, "Average Delivery Time"], number_format)
        worksheet.write(row_num, 2, summary_df.at[0, "Fast Deliveries (≤ 45 mins)"], int_format)
        worksheet.write(row_num, 3, summary_df.at[0, "High Ratings (≥ 4.5)"], int_format)