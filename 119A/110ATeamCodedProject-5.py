#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Sat Apr 25 12:17:36 2020

@author: luistovar
"""

import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
x1 = pd.ExcelFile("OSSalesData.xlsx")
lines = "="*33 + "\n"
SalesData = x1.parse("Orders")

# =============================================================================
# #1
# Professor's original code. We are using it as a reference point.
# Shows profits per Sub-Category
# =============================================================================

def SubCatProfits():
    SubCatData = SalesData[["Sub-Category", "Profit"]]
    print(lines)
    SubCatProfit = SubCatData.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
    print(SubCatProfit.head(10))


# =============================================================================
# #2
# The output of this function will display three items regarding yearly profits throughout the 4 years in the Orders file.
# It is showing which years were the most profitabl
# The first item is a list of years and the total profit earned in each of those years.
# The second output shows a pie chart regarding the years and profits
# The third output is a line plot that shows the change in profits throghout the years
# =============================================================================
    
def AnnualProfitChart():
    SalesDataYear = SalesData
    SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
    YearlyProfits = SalesDataYear[["Year", "Profit"]]
    YearlyProfitTotal = YearlyProfits.groupby("Year").sum()  
    print(lines)
    print(YearlyProfitTotal) #displays list of years and profits

    #Display pie chart
    YearlyProfitTotal = YearlyProfitTotal.reset_index()
    
    plt.figure(figsize=(8,8))
    plt.title("Yearly Profit Percentages")
    plt.pie(YearlyProfitTotal.Profit, labels = YearlyProfitTotal.Year, autopct = "%1.1f%%", counterclock = False)
    plt.show()
    
    #Display line Plot
    lineplot1 = sns.lineplot(x = "Year", y = "Profit", data = YearlyProfitTotal)
    lineplot1.set_title("Profits by Year")


# =============================================================================
# #3
# This function will provide the profit for each product name in the 5 lowest performing sub categories, 
# the profit ranges from the least amount of profit to the most 
# =============================================================================
    
def ProfitSubCat():
    ProductData = SalesData[["Sub-Category", "Product Name","Profit"]]
    ProfSubCats = ["Tables", "Bookcases", "Supplies", "Fasteners", "Machines"]
    
    for profits in ProfSubCats:
        print(lines)
        ProductInfo = ProductData.loc[ProductData["Sub-Category"]==profits]
        SubCatProfit = ProductInfo.groupby(by = "Product Name").sum().sort_values(by = "Profit")
        print(profits)
        pd.options.display.float_format = '$ {:.2f}'.format
        print(SubCatProfit * 100)    


# =============================================================================
# #4
# This function will provie us with the quantity of products sold in each segment based off of the 4 regions       
# =============================================================================
        
def QuantityByRegion():
    Regions = SalesData.Region.unique()
    print(Regions)
    SubCatData = SalesData[["Segment", "Quantity", "Region",]]

    for region in Regions:
        RegionSubCatData = SubCatData.loc[SubCatData["Region"]==region]
        SubCatProfit = RegionSubCatData.groupby(by = "Segment").sum().sort_values(by="Quantity")
        
        print(lines)
        print(region)
        
        print(SubCatProfit.head(10))


# =============================================================================
# #5
# This function will show the top 10 lowest profiting products in each state that has a buyer in it. 
# For example if a state only has 2 products, it will only show 2. 
# If it has 21, it will show the first 10 products and the amount of profit made along with the state's name.
# =============================================================================
        
def StateSubCatProf():
    states = SalesData.State.unique()
    print(states)
    SubCatData = SalesData[["Product Name", "Profit", "State"]]

    for state in states:
        StateSubCatData = SubCatData.loc[SubCatData["State"]==state]
        SubCatProfit = StateSubCatData.groupby(by = "Product Name").sum().sort_values(by="Profit")
        print(lines)
        print(state)
        print(SubCatProfit.head(10))


# =============================================================================
# #6
# This function will show the total amount of products bought by each segment. 
# It will split into the three segments and show the product name and amount bought in the respective segment.        
# =============================================================================

def SegAmt():
    segments = SalesData.Segment.unique()
    print(segments)
    ProdData = SalesData[["Product Name", "Segment", "Quantity"]]
    for segment in segments:
        RegProdData = ProdData.loc[ProdData["Segment"]==segment]
        ProdQuantity = RegProdData.groupby(by = "Product Name").sum().sort_values(by = "Quantity")
        print(lines)
        print(segment)
        print(ProdQuantity)


# =============================================================================
# #7
# Original function provided by professor.
# Shows annual profit per Sub-Category
# =============================================================================
        
def AnnualEfProf():
    SalesDataYear = SalesData
    SalesDataYear["Year"]=SalesDataYear["Order Date"].dt.year
    years = SalesDataYear.Year.unique()
    print(years)
    
    SubCatData=SalesDataYear[["Sub-Category", "Profit", "Year"]]
    
    for year in years:
        SubCatDataByYear = SubCatData.loc[SubCatData["Year"]==year]
        SubCatProfitNoYear = SubCatDataByYear[["Sub-Category", "Profit"]]
        SubCatProfit = SubCatProfitNoYear.groupby(by = "Sub-Category").sum().sort_values(by = "Profit")
        print(lines)
        print(year)
        pd.options.display.float_format = '$ {:.2f}'.format
        print(SubCatProfit.head(10))


# =============================================================================
# #8
# This function will show the discounts associated with the products in the top 5 underperforming/low profits Sub-Categories. 
# I derived the Sub-Categories by running the first function which shows the Sub-Categories ascending by profit. 
# The first 5 show the least profit. 
# The output will show the Sub-Category name, followed by the product name in each category and the discount associated with that product.
# =============================================================================

def DiscountSubCat():
    ProductData = SalesData[["Sub-Category", "Product Name", "Discount"]]
    DiscSubCats = ["Tables", "Bookcases", "Supplies", "Fasteners", "Machines"]
    
    for discsubcat in DiscSubCats:
        print(lines)
        ProductInfo = ProductData.loc[ProductData["Sub-Category"]== discsubcat]
        SubCatDiscount = ProductInfo.groupby(by = "Product Name").mean().sort_values(by = "Discount")
        print(discsubcat)
        pd.options.display.float_format = '{:.2f}%'.format
        print(SubCatDiscount * 100)


# =============================================================================
# #9
# The next three fucnitons are oriented towards customer loyalty. 
# After initial execution, the customer will be able to choose between three customer loyalty insights
# =============================================================================

# =============================================================================
# The output will display the names of the top 5% of OS's frequent customers
# =============================================================================
def MostFreqCustomers():
    MostFrequent = SalesData["Customer Name"].value_counts()
    total_customers = MostFrequent.shape
    TopPercent = int(total_customers[0]* .05) # .05 is the percent value. Can change between 0 - 1 to see other percentages.
    print(MostFrequent.head(TopPercent))

# =============================================================================
# The output will display the names of the top 5% of OS's recent customers
# =============================================================================
def RecentCustomers():
    Customer_OrderDate = SalesData[["Customer Name","Order Date"]]
    SortDates = Customer_OrderDate.sort_values(by="Order Date", ascending =False)
    SortDates_NoDup = SortDates.drop_duplicates()
    total_customers = SortDates_NoDup.shape
    TopPercent = int(total_customers[0]*.05) # .05 is the percent value. Can change between 0 - 1 to see other percentages.
    print(SortDates_NoDup.head(TopPercent))
    
# =============================================================================
# The output would show the names and sales of the top 10% of OS's highest spending customers
# =============================================================================   
def HighestSpendingCustomers():
    Customer_Profit= SalesData[["Customer Name", "Sales"]]
    Customer_Total_Profit = Customer_Profit.groupby(by="Customer Name").sum().sort_values(by="Sales", ascending =False).reset_index()
    total_customers = Customer_Total_Profit.shape
    TopPercent = int(total_customers[0]*.10) # .10 is the percent value. Can change between 0 - 1 to see other percentages.
    print(Customer_Total_Profit.head(TopPercent))


# =============================================================================
# #10
# This function will show the year to year sales from 2016 to 2019 for each category
# Each category shown is added up together into total sales
# The 4 regions show the same sub-categories, however the discount rate varies from region to region
# =============================================================================

def YearlyCatSales():
    SalesDataYear = SalesData
    SalesDataYear["Year"] = SalesDataYear["Order Date"].dt.year
    years = SalesDataYear.Year.unique()
    CatData = SalesDataYear[["Category", "Sales", "Year"]]
    print(years)
    
    for year in years:
        CatDataByYear = CatData.loc[CatData["Year"]==year]
        CatProfitNoYear = CatDataByYear[["Category", "Sales"]]
        CatProfit = CatProfitNoYear.groupby(by = "Category").sum().sort_values(by = "Sales")
        print(lines)
        print(year)
        pd.options.display.float_format = '$ {:.2f}'.format
        print(CatProfit.head(10))
    
    
# =============================================================================
# #11
# This function will display discounts per each region
# The output would allow the viewer to see each sub-category, region, and the discount for the sub-category
# =============================================================================
        
def DiscountPerRegion():
    regions = SalesData.Region.unique()
    SubCatDiscountRegData = SalesData[["Sub-Category", "Discount", "Region"]]

    for region in regions:
        SubCatInfo = SubCatDiscountRegData.loc[SubCatDiscountRegData["Region"]==region]
        SubCatDiscount = SubCatInfo.groupby(by="Sub-Category").mean().sort_values(by="Discount")
        print(lines)
        print(region)
        pd.options.display.float_format = '{:.2f}%'.format
        print(SubCatDiscount.head(10) * 100)
        


# =============================================================================
# #Main Menu function that displays user commands
# =============================================================================
        
def MainMenu():
    print("\n" + "*" * 90)
    print("\n Enter -1- to see Sub-Category Profits" +
          "\n Enter -2- to see visuals representing the yearly distribution of OS's profits" +
          "\n Enter -3- to see the profit of each product in the top 5 underperforming Sub-Categories" +
          "\n Enter -4- to see the quantaties of products sold in each region based off of segments" +
          "\n Enter -5- to see all products by state that are producing the least profits" +
          "\n Enter -6- to see all underperforming products and quantities bought by each segment" +
          "\n Enter -7- to see all profits of products by year" +
          "\n Enter -8- to see all discounts applied to products of top 5 underperforming Sub-Category" +
          "\n Enter -9- to see insights about OS's loyal customers" +
          "\n Enter -10- to see the yearly sales of each Category" +
          "\n Enter -11- to see the average discount per Sub-Category" + 
          "\n Enter -12- to exit")
    print("\n" + "*" * 90)
    choice = input("Please enter a number 1 - 12: ") 
    if choice == "1":
        SubCatProfits()
        MainMenu()
    elif choice == "2":
        AnnualProfitChart()
        MainMenu()    
    elif choice == "3":
        ProfitSubCat()
        MainMenu()
    elif choice == "4":
        QuantityByRegion()
        MainMenu()
    elif choice == "5":
        StateSubCatProf()
        MainMenu()
    elif choice == "6":
        SegAmt()
        MainMenu()
    elif choice == "7":
        AnnualEfProf()
        MainMenu()
    elif choice == "8":
        DiscountSubCat()
        MainMenu()
    elif choice == "9":
        print("\n" + "*" * 65)
        print("\n Enter 9.1 to see the top 5% of OS's frequent customers" +
              "\n Enter 9.2 to see the top 5% of OS's recent customers" +
              "\n Enter 9.3 to see the top 10% of OS's highest spending customers")
        print("\n" + "*" * 65)
        choice1 = input("Please enter 9.1, 9.2, or 9.3 to see desired insights: ")
        if choice1 == "9.1":
            MostFreqCustomers()
            MainMenu()
        elif choice1 == "9.2":
            RecentCustomers()
            MainMenu()
        elif choice1 == "9.3":
            HighestSpendingCustomers()
            MainMenu()
        else:
            print("Invalid input, please try again")
            MainMenu()
        MainMenu()
    elif choice == "10":
        YearlyCatSales()
        MainMenu()
    elif choice == "11":
        DiscountPerRegion()
        MainMenu()
    elif choice == "12":
        exit()
    else:
        print("Invalid input, please try again")
        MainMenu()      
print("\nWelcome to Office Solutions Data Analytics System")
MainMenu()
