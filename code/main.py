  # Importing necessary libraries
import pandas as pd
import matplotlib.pyplot as pl
import numpy as np
import seaborn as sns

def main():

    #Loading the data file
    sales=pd.read_excel(r'D:\Swayam\technical_skills\internship projects\Amazon sales analysis\data\Amazon sales.xlsx')
    print(sales.head()) #printing first 5 rows, to get insight on type of data
    sns.set_style("whitegrid")
    region_wise_profit=sales.groupby('Region')['Total Profit'].mean()
    region_wise_profit=pd.DataFrame(region_wise_profit)
    region_wise_revenue=sales.groupby('Region')['Total Revenue'].mean()
    region_wise_revenue=pd.DataFrame(region_wise_revenue)
    region_wise_cost=sales.groupby('Region')['Total Cost'].mean()
    region_wise_cost=pd.DataFrame(region_wise_cost)
    region_wise_unitsold=sales.groupby('Region')['Units Sold'].mean()
    region_wise_unitsold=pd.DataFrame(region_wise_unitsold)

    
    
    # converting the appropriate columns to datetime 
    sales['Order Date'] = pd.to_datetime(sales['Order Date'])
    sales['Ship Date']=pd.to_datetime(sales['Ship Date'])






    unique_country=sales['Country'].unique()
    unique_region=sales['Region'].unique()

    
    def research_question():
        sorted_values_profit = region_wise_profit['Total Profit'].sort_values(ascending=False)

        #Evaluating the trend
        pl.figure(figsize=(10,10))
        ax=sns.barplot(x='Region',y='Total Profit',data=region_wise_profit, order=sorted_values_profit.index)
        ax.set_xticklabels(ax.get_xticklabels(), fontsize=8)    #SETTING FONT SIZE =7
        pl.tight_layout()
        pl.show()

        # Research Question: What are key Factors influencing the profitability in a region and plausible solutions to maximising it. 
    research_question()

    def EDA():
        
        pl.figure(figsize=(20,20))
        pl.subplot(2,2,1)
        sorted_values_revenue = region_wise_revenue['Total Revenue'].sort_values(ascending=False)
        ax=sns.barplot(x='Region',y='Total Revenue',data=region_wise_revenue,order=sorted_values_revenue.index)
        ax.set_xticklabels(ax.get_xticklabels(), fontsize=7,rotation=25, ha="right")    #SETTING FONT SIZE =7
        pl.subplot(2,2,2)
        sorted_values_profit = region_wise_profit['Total Profit'].sort_values(ascending=False)
        
        ax=sns.barplot(x='Region',y='Total Profit',data=region_wise_profit,order=sorted_values_profit.index)
        ax.set_xticklabels(ax.get_xticklabels(), fontsize=7,rotation=25, ha="right")    #SETTING FONT SIZE =7

        pl.subplot(2,2,3)
        sorted_values_cost=region_wise_cost['Total Cost'].sort_values(ascending=False)
        ax=sns.barplot(x='Region',y='Total Cost',data=region_wise_cost,order=sorted_values_cost.index)
        ax.set_xticklabels(ax.get_xticklabels(), fontsize=7,rotation=25, ha="right")    #SETTING FONT SIZE =7

        pl.subplot(2,2,4)
        sorted_values_unitsold=region_wise_unitsold['Units Sold'].sort_values(ascending=False)
        ax=sns.barplot(x='Region',y='Units Sold',data=region_wise_unitsold,order=sorted_values_unitsold.index)
        ax.set_xticklabels(ax.get_xticklabels(), fontsize=7,rotation=25, ha="right")    #SETTING FONT SIZE =7

        pl.tight_layout()
        pl.show()

        country_in_region=sales.groupby('Region')['Country'].nunique()

        item_type_sales=sales.groupby(['Region','Item Type'])['Units Sold'].sum()
        item_type_sales=pd.DataFrame(item_type_sales)
        item_type_sales.to_excel('item_type_sales.xlsx')
       
        item_type_sales.reset_index(inplace=True)


        pivot_table = item_type_sales.pivot_table(index='Region', columns='Item Type', values='Units Sold', fill_value=0)


        
   


    EDA()

        





main()