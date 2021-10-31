import pandas as pd
import Business_Enablement
import CEO_Office
import Delivery
import Marketing_Growth
import Strategic_Initiatives
import Product
# Import new Business Unit if any are added

df_BE = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Business Enablement.xlsx")
df_Marketing_Growth = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Marketing & Growth.xlsx")
df_Delivery = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Delivery.xlsx")
df_CEO_Office = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\CEO Office.xlsx")
df_Product = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Product.xlsx")
df_Strategic_Initiatives = pd.read_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Strategic Initiatives.xlsx")

df_merge = pd.concat([df_BE, df_Marketing_Growth, df_Delivery, df_CEO_Office, df_Product,
                      df_Strategic_Initiatives])

df_merge.reset_index(inplace=True)
df_merge = df_merge.iloc[:, 2:13]
df_merge.to_excel(r"C:\Users\Diwaker Sinha\Desktop\MPL\Combined.xlsx")
