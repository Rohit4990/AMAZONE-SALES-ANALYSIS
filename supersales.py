import pandas as pd

df = pd.read_excel("sample supersales.xlsx")
pd.set_option('display.max_rows',None)
pd.set_option('display.max_columns',None)
pd.set_option('display.width',None)
pd.set_option('display.max_colwidth',None)
df.columns = df.columns.str.strip()

print(df)



print(df.head(10))
print(df.tail(10))
print(df.shape)

df = df.drop_duplicates(subset=["Order ID","Product Name","Sales","Quantity","Profit"])
print(df)


print(df.info())

df["Order Date"] = pd.to_datetime(df["Order Date"])
print(df)

df["Ship Date"] = pd.to_datetime(df["Ship Date"])
print(df)


print(df.info())
print(df.shape)

print(df.columns.tolist())

#rename columns
df.rename(columns={

    "Order ID":"Order_ID",
    "Product Name":"Product_Name",
    "Order Date":"Order_Date",
    "Ship Date":"Ship_Date",


},inplace=True)

print(df)

print(df.columns)


print(df["Sales"])

q1 = df["Sales"].quantile(0.25)
q3 = df["Sales"].quantile(0.75)
iqr = q3 - q1

lower_limit = q1-1.5*iqr
upper_limit = q3+1.5*iqr

outliers = df[(df["Sales"] < lower_limit) | (df["Sales"] > upper_limit)]
print(outliers[["Sales"]])
print(outliers.shape)

print(df.info())



import matplotlib.pyplot as plt
plt.figure(figsize = (10,10))
monthly_sales = df.groupby(df["Order_Date"].dt.to_period("M"))["Sales"].sum()
monthly_sales.index = monthly_sales.index.astype(str)
plt.plot(monthly_sales.index,monthly_sales.values,color="gold",marker="o")
plt.xticks(rotation=70)
plt.title("Monthly Sales",fontsize=14,fontweight="bold",color="gold")
plt.xlabel("Month",fontsize=10,fontweight="bold",color="gold")
plt.ylabel("Sales",fontsize=10,fontweight="bold",color="gold")
plt.tight_layout()
plt.savefig("Monthly Sales",dpi=300)
plt.show()

monthly_sales = df.groupby(df["Order_Date"].dt.to_period("Y"))["Sales"].sum()
monthly_sales.index = monthly_sales.index.astype(str)
plt.plot(monthly_sales.index,monthly_sales.values,color="pink",marker="o")
plt.title("YEARLY SALES ANALYSIS",fontsize=10,fontweight="bold")
plt.ylabel("TOTAL_SALES",fontsize=10,fontweight="bold")
plt.xlabel("YEAR",fontsize=10,fontweight="bold")
plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.savefig("yearly sales ANALYSIS",dpi=300)
plt.show()

# Step 1: Clean Product Name column
df['Product_Name'] = df['Product_Name'].str.strip().str.title()  # remove spaces + proper case

top_products = df.groupby("Product_Name")["Sales"].sum().sort_values(ascending=False).head(10)
print(top_products)

plt.bar(top_products.index,top_products.values,color="orange",label="TOP PRODUCTS")
plt.title("TOP 10 PRODUCTS",fontsize=10,fontweight="bold")
plt.xlabel("Product Name",fontsize=10,fontweight="bold")
plt.ylabel("TOP_SALES",fontsize=10,fontweight="bold")
plt.legend(loc="upper center")
plt.xticks(rotation=65)
plt.tight_layout()
plt.savefig("TOP 10 PRODUCTS",dpi=300)
plt.show()


sub = df.groupby("Sub-Category")["Sales"].sum()

plt.bar(sub.index,sub.values,color="gold",label="SUB CATEGORY")
plt.title("SUB CATEGORY WISE SALES",fontsize=10,fontweight="bold")
plt.xlabel("Sub Category",fontsize=10,fontweight="bold")
plt.ylabel("TOTAL_SALES",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")
plt.xticks(rotation=50)
plt.tight_layout()
plt.show()


cat = df.groupby("Category")["Sales"].sum()

plt.pie(cat.values,labels=cat.index,autopct="%1.1f%%")
plt.title("CATEGORY WISE TOTAL SALES",fontsize=10,fontweight="bold")
plt.tight_layout()
plt.show()


region = df.groupby("Region")["Sales"].sum().sort_values(ascending=False)

plt.barh(region.index,region.values,color="skyblue",label="REGION")
plt.title("REGION WISE SALES",fontsize=10,fontweight="bold")
plt.xlabel("TOTAL_SALES",fontsize=10,fontweight="bold")
plt.ylabel("REGION",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")
plt.tight_layout()
plt.show()



#subplot
plt.figure(figsize=(20,12))

plt.subplot(2,1,1)
plt.bar(sub.index,sub.values,color="gold",label="SUB CATEGORY")
plt.title("SUB CATEGORY WISE SALES",fontsize=10,fontweight="bold")
plt.xlabel("Sub Category",fontsize=10,fontweight="bold")
plt.ylabel("TOTAL_SALES",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")
plt.xticks(rotation=50)

plt.subplot(2,1,2)

profit = df.groupby("Sub-Category")["Profit"].sum()
plt.bar(profit.index,profit.values,color="purple",label="PROFIT")
plt.title("SUB CATEGORY WISE PROFIT",fontsize=10,fontweight="bold")
plt.xlabel("Sub Category",fontsize=10,fontweight="bold")
plt.ylabel("TOTAL_PROFIT",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.suptitle("PROFIT VS SALES BY SUB CATEGORY",fontsize=14,fontweight="bold")
plt.savefig("SUB CATEGORY WISE PROFIT",dpi=300)

plt.show()


plt.figure(figsize=(20,12))
plt.subplot(1,2,1)
plt.barh(region.index,region.values,color="skyblue",label="REGION")
plt.title("REGION WISE SALES",fontsize=10,fontweight="bold")
plt.xlabel("TOTAL_SALES",fontsize=10,fontweight="bold")
plt.ylabel("REGION",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")

region = df.groupby("Region")["Profit"].sum().sort_values(ascending=False)

plt.subplot(1,2,2)
plt.barh(region.index,region.values,color="lightgreen",label="REGION")
plt.title("REGION WISE PROFIT",fontsize=10,fontweight="bold")
plt.xlabel("TOTAL_PROFIT",fontsize=10,fontweight="bold")
plt.ylabel("REGION",fontsize=10,fontweight="bold")
plt.legend(loc="upper right")

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.suptitle("PROFIT VS SALES BY REGION",fontsize=14,fontweight="bold")
plt.savefig("REGION WISE PROFIT",dpi=300)

plt.show()


plt.figure(figsize=(20,12))
plt.subplot(1,2,1)
plt.pie(cat.values,labels=cat.index,autopct="%1.1f%%")
plt.title("CATEGORY WISE TOTAL SALES",fontsize=12,fontweight="bold")

profit = df.groupby("Category")["Profit"].sum()
plt.subplot(1,2,2)
plt.pie(profit.values,labels=profit.index,autopct="%1.1f%%",colors=["gold","skyblue","pink"])
plt.title("CATEGORY WISE PROFIT",fontsize=12,fontweight="bold")

plt.tight_layout(rect=[0, 0, 1, 0.95])
plt.suptitle("PROFIT VS SALES BY CATEGORY",fontsize=14,fontweight="bold")
plt.savefig("CATEGORY WISE PROFIT",dpi=300)

plt.show()


df.to_csv("supersales10.csv",index=False)
print(df)
