import pandas as pd
df = pd.read_excel ("C://Users//ADMIN//Desktop//PYTHON FILES//Assignment_Retail_Sales_Analysis_Data.xlsx")
df
df.isnull().sum()
df["OrderDate"] =pd.to_datetime(
    df["OrderDate"], format="mixed"
)
df
df["OrderDate"].unique()
df["Age"] = df["Age"].replace([0,150,120],pd.NA)
df
df["Age"].unique()
df.loc[df["Age"] < 0,"Age"]=pd.NA
df
df["Age"].mean()
df["Age"]=df["Age"].fillna(df["Age"].mean()).astype(int)
df
df["Product"].unique()
df["Product"]=df["Product"].fillna(df["Product"].mode()[0])
df
df["UnitPrice"]=df["UnitPrice"].fillna(df["UnitPrice"].mean()).round(2)
df
df["Discount"]= df["Discount"].fillna(df["Discount"].mean())
df
df["Sales"].mean()
df.loc[df["Sales"] < 0,"Sales"]=pd.NA
df
df["Sales"]=df["Sales"].fillna(df["Sales"].mean()).round(2)
df
df["Sales"].unique()
df["Cost"].mean()
df["Cost"]=df["Cost"].fillna(df["Cost"].mean()).round(2)
df
df["Cost"].unique()
df["Profit"]=df["Profit"].fillna(df["Profit"].mean()).round(2)
df
df["PaymentMode"]=df["PaymentMode"].fillna(df["PaymentMode"].mode()[0])
df
df["PaymentMode"].unique()
df["PaymentMode"] =df["PaymentMode"].map({
          "UPI":"UPI",
          "upi":"UPI",
          "GPay" : "GooglePay",
          "Cash" : "Cash",
          "GooglePay" :"GooglePay",
          "Card" :"Card"
})
df
df["Rating"]=df["Rating"].fillna(df["Rating"].mean()).astype(int)
df
df["Rating"].unique()
df["DeliveryDays"] = df["DeliveryDays"].replace([45],pd.NA)
df["DeliveryDays"] = df["DeliveryDays"].fillna(df["DeliveryDays"].mean()).round(1).astype(int)
df
df["DeliveryDays"].unique()
df["Returned"].unique()
df["Returned"] = df["Returned"].fillna(df["Returned"].mode()[0])
df
df["Quantity"].unique()
df["Quantity"] = df["Quantity"].replace(['abc',-2,-1],pd.NA)
df
df["Quantity"]=df["Quantity"].fillna(df["Quantity"].mean()).astype(int)
df
df["State"].unique()
df["State"] =df["State"].map({
          'West Bengal' : 'West Bengal', 
          'Telangana' : 'Telangana', 
          'Delhi':'Delhi', 
          'Maharashtra':'Maharashtra',
          'Karnataka':'Karnataka',
          'TG':'Telangana', 
          'Gujarat':'Gujarat', 
          'Haryana':'Haryana', 
          'Tamil Nadu':'Tamilnadu',
          'Assam': 'Assam', 
          'Goa':'Goa', 
          'Kerala':'Kerala',
          'Punjab':'Punjab',
          'Odisha' :'Odisha', 
          'Telengana':'Telangana',
          'TELANGANA':'Telangana'
})
df
df["City"] = df["City"].map({
    'Kolkata': 'Kolkata',
    'Calcutta': 'Kolkata',
    'HYD': 'Hyderabad',
    'Hyd': 'Hyderabad',
    'Hyderabad': 'Hyderabad',
    'hyderabad': 'Hyderabad',
    ' HYD ': 'Hyderabad',
    'Delhi': 'Delhi',
    'New Delhi': 'Delhi',
    'Bombay': 'Mumbai',
    'Mumbai': 'Mumbai',
    'BLR': 'Bangalore',
    'Bangalore': 'Bangalore',
    'Bengaluru': 'Bangalore',
    'Pune': 'Pune',
    'Ahmedabad': 'Ahmedabad',
    'Gurgaon': 'Gurgaon',
    'Chennai': 'Chennai',
    'Madras': 'Chennai',
    'Guwahati': 'Guwahati',
    'Panaji': 'Panaji',
    'Panjim': 'Panaji',
    'Kochi': 'Cochin',
    'Cochin': 'Cochin',
    'Surat': 'Surat',
    'Amritsar': 'Amritsar',
    'Bhubaneswar': 'Bhubaneswar',
    'BBSR': 'Bhubaneswar'
})
df
df["City"].unique()
df["City"] = df["City"].str.strip()
df
df["City"].value_counts()
df["Gender"] = df["Gender"].map({
    "Female":"Female",
    "female":"Female",
    "Male":"Male",
    "male":"Male",
    "F":"Female",
    "M":"Male"
    })
df
df["Gender"].unique()
df["CustomerID"].unique()
df
df["Online/Offline"] = df["PaymentMode"].map({
    "UPI":"Online",
    "GooglePay":"Online",
    "Card":"offline",
    "Cash":"Offline"
    
    })
df
df= df.rename(columns= {"Quantity":"Units_Sold"})
df
df= df.rename(columns= {"Sales":"Price_per_Unit"})
df
df =df.rename(columns= {"UnitPrice":"Selling_Price"})
df
df =df.rename(columns= {"Discount":"Discount_%"})
df   
df=df.rename(columns= {"DeliveryDays":"Delivery_Days"})
df
df=df.rename(columns= {"Returned":"Return_Status"})
df
df=df.rename(columns= {"OrderID":"Order_ID"})
df
df=df.rename(columns= {"OrderDate":"Order_Date"})
df
df=df.rename(columns= {"CustomerID":"C_ID"})
df
df=df.rename(columns= {"CustomerName":"C_Name"}) 
df
df=df.rename(columns= {"SalesRep":"   Sales_person"})   
df
df["Year"] =df["Order_Date"].dt.year
df
df["Month"] =df["Order_Date"].dt.month
df
df["Day"] =df["Order_Date"].dt.day
df
df["Day_FullName"] =df["Order_Date"].dt.day_name()
df
df["Day_ShortName"] =df["Order_Date"].dt.strftime("%a")
df
df["Month_FullName"] =df["Order_Date"].dt.month_name()
df
df["Month_ShortName"] =df["Order_Date"].dt.strftime("%b")
df
df.info()
df["Order_Date"]=pd.to_datetime(df["Order_Date"],format="mixed")
df["Order_Date"].dtype
df
df[df["Order_ID"]==1099]
df =df.drop_duplicates()
df
df["Order_ID"].unique()
df["Total_Sales"] =df["Units_Sold"]*df["Price_per_Unit"]
df
path = r"C:\\Users\\ADMIN\\Desktop\\PYTHON FILES\\Python_Assignment_Data.xlsx"
df.to_excel(path, index = False)
