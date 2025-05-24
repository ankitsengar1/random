import pandas as pd
import numpy as np
from pptx.enum.dml import MSO_COLOR_TYPE
from pptx.enum.dml import MSO_THEME_COLOR_INDEX
from pptx.dml.color import RGBColor
import warnings
from formatting_functions import format_with_locale
import variables
warnings.simplefilter(action='ignore', category=FutureWarning)

def edit_table(slide,df):
    for shape in slide.shapes:
        if shape.has_table:
            table = shape.table
            for col in range(0,len(table.columns)):
                for row in range(0,len(table.rows)):
                    for paragraph in table.cell(row,col).text_frame.paragraphs:
                        first_run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
                        font = first_run.font
                        font_name = font.name
                        font_size = font.size
                        font_bold = font.bold
                        font_italic = font.italic
                        font_underline = font.underline
                        text = paragraph.text
                        if font.color.type == MSO_COLOR_TYPE.SCHEME:
                            # Get RGB from theme color
                            theme_color = font.color.theme_color
                            font_rgb_flag = 1
                        elif font.color.type == None:
                            font_rgb_flag = 0
                        elif font.color.type == MSO_COLOR_TYPE.PRESET:
                            color_index = font.color.theme_color
                            font_rgb_flag = 3
                        else:
                            # Already RGB color, use it directly
                            rgb = font.color.rgb
                            font_rgb_flag = 2
                
                        # if row==0:
                        #     pass
                        # else:
                        paragraph.clear()                    
                        new_run = paragraph.add_run()
                        new_run.font.name = font_name
                        new_run.font.size = font_size
                        new_run.font.bold = font_bold
                        new_run.font.underline = font_underline
                        new_run.font.italic = font_italic
                        # print(font_rgb_flag)
                        if font_rgb_flag == 1:
                            new_run.font.color.theme_color = theme_color
                        elif font_rgb_flag == 2:
                            new_run.font.color.rgb = rgb
                        elif font_rgb_flag == 3:
                            new_run.font.color.color_name = color_name = MSO_THEME_COLOR_INDEX(color_index).name

                        if row > 0: # leaving headers values
                            value = str(df.iloc[row-1,col])

                            if col in [4,5]:
                                old_value = str(df.iloc[row-1,col])
                                value = float(value.strip('%'))

                                if value < 0:
                                    new_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
                                elif value > 0:
                                    new_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
                                else:
                                    new_run.font.color.rgb = RGBColor(0, 0, 0)  # Green
                                
                                new_run.text = old_value

                            else:
                                new_run.text = str(value)

                        elif row == 0: # headers
                            new_run.text = str(list(df.columns.values)[col])
                        
    return slide

def create_df(chogori_ppt_xl_path,sheet_name):

    df = pd.read_excel(chogori_ppt_xl_path,sheet_name = sheet_name)
    df.rename(columns={df.columns[0]:"KPIS",df.columns[1]:variables.curr_month_prev_yr,df.columns[2]:variables.prev_month,df.columns[3]:variables.current_month},inplace = True)
    replacements = {'enrollments':'Enrollments', 'transacting_customers': 'Transacting Customers','sales':'Total Sales',
                    'bills':'Total Bills','abv':'ABV(Overall)','onetimer_abv':'New one timer ABV','new_repeat_abv':'New repeater ABV','old_repeat_abv':'Old repeater ABV','amv':'AMV(Overall)',
                    'onetimer_amv':'New one timer AMV','new_repeat_amv':'New repeater AMV','old_repeat_amv':'Old repeater AMV',"transaction_points_issued":"Transaction Points Issued","points_redeemed":"Points Redeemed"}
    df["KPIS"] = df["KPIS"].replace(replacements)

    new_onetimer = (df[df["KPIS"] == "onetimer"].iloc[:, 1:].values / df[df["KPIS"] == "Transacting Customers"].iloc[:, 1:].values)*100
    new_onetimer_row = pd.DataFrame([["New one timer%"] + new_onetimer.flatten().tolist()], columns=df.columns)

    new_repeater = (df[df["KPIS"] == "new_repeater"].iloc[:, 1:].values / df[df["KPIS"] == "Transacting Customers"].iloc[:, 1:].values)*100
    new_repeater_row = pd.DataFrame([["New Repeater%"] + new_repeater.flatten().tolist()], columns=df.columns)

    old_repeater = (df[df["KPIS"] == "old_repeater"].iloc[:, 1:].values / df[df["KPIS"] == "Transacting Customers"].iloc[:, 1:].values)*100
    old_repeater_row = pd.DataFrame([["Old Repeater%"] + old_repeater.flatten().tolist()], columns=df.columns)

    new_onetimer_sales = (df[df["KPIS"] == "onetimer_sales"].iloc[:, 1:].values / df[df["KPIS"] == "Total Sales"].iloc[:, 1:].values)*100
    new_onetimer_sales_row = pd.DataFrame([["New one timer Sales%"] + new_onetimer_sales.flatten().tolist()], columns=df.columns)

    new_repeater_sales = (df[df["KPIS"] == "new_repeat_sales"].iloc[:, 1:].values / df[df["KPIS"] == "Total Sales"].iloc[:, 1:].values)*100
    new_repeater_sales_row = pd.DataFrame([["New Repeater Sales%"] + new_repeater_sales.flatten().tolist()], columns=df.columns)

    old_repeater_sales = (df[df["KPIS"] == "old_repeat_sales"].iloc[:, 1:].values / df[df["KPIS"] == "Total Sales"].iloc[:, 1:].values)*100
    old_repeater_sales_row = pd.DataFrame([["Old Repeater Sales%"] + old_repeater_sales.flatten().tolist()], columns=df.columns)

    new_onetimer_bills = (df[df["KPIS"] == "onetimer_bills"].iloc[:, 1:].values / df[df["KPIS"] == "Total Bills"].iloc[:, 1:].values)*100
    new_onetimer_bills_row = pd.DataFrame([["New one timer Bills%"] + new_onetimer_bills.flatten().tolist()], columns=df.columns)

    new_repeater_bills = (df[df["KPIS"] == "new_repeat_bills"].iloc[:, 1:].values / df[df["KPIS"] == "Total Bills"].iloc[:, 1:].values)*100
    new_repeater_bills_row = pd.DataFrame([["New Repeater Bills%"] + new_repeater_bills.flatten().tolist()], columns=df.columns)

    old_repeater_bills = (df[df["KPIS"] == "old_repeat_bills"].iloc[:, 1:].values / df[df["KPIS"] == "Total Bills"].iloc[:, 1:].values)*100
    old_repeater_bills_row = pd.DataFrame([["Old Repeater Bills%"] + old_repeater_bills.flatten().tolist()], columns=df.columns)

    df = pd.concat([df,new_onetimer_row,new_repeater_row,old_repeater_row,new_onetimer_sales_row,new_repeater_sales_row,old_repeater_sales_row,new_onetimer_bills_row,new_repeater_bills_row,old_repeater_bills_row], ignore_index=True)
    
    df = df[~df["KPIS"].isin(["onetimer","new_repeater","old_repeater","onetimer_sales","new_repeat_sales","old_repeat_sales","onetimer_bills","new_repeat_bills","old_repeat_bills"])]
    df.reset_index(drop=True, inplace=True)

    percent_kpis = ["New one timer%", "New Repeater%", "Old Repeater%", "New one timer Sales%", "New Repeater Sales%","Old Repeater Sales%","New one timer Bills%","New Repeater Bills%","Old Repeater Bills%"]

    df1 = df[~df["KPIS"].isin(percent_kpis)].reset_index(drop=True)
    df2 = df[df["KPIS"].isin(percent_kpis)].reset_index(drop=True)

    df1['% Change-MoM'] = (df1[df1.columns[3]]-df1[df1.columns[2]])/df1[df1.columns[2]]*100
    df1['% Change-YoY'] = (df1[df1.columns[3]]-df1[df1.columns[1]])/df1[df1.columns[1]]*100
    df1['% Change-MoM'] = df1['% Change-MoM'].round(1).astype(str) + '%'
    df1['% Change-YoY'] = df1['% Change-YoY'].round(1).astype(str) + '%'
    df1['% Change-MoM'] = df1['% Change-MoM'].replace('inf%', "0.0%")
    df1['% Change-YoY'] = df1['% Change-YoY'].replace('inf%', "0.0%")

    df2['% Change-MoM'] = (df2[df2.columns[3]]-df2[df2.columns[2]])
    df2['% Change-YoY'] = (df2[df2.columns[3]]-df2[df2.columns[1]])

    df2['% Change-MoM'] = df2['% Change-MoM'].round(1).astype(str) + '%'
    df2['% Change-YoY'] = df2['% Change-YoY'].round(1).astype(str) + '%'

    df  = pd.concat([df1, df2], ignore_index=True)

    exceptions = ["New one timer%", "New Repeater%", "Old Repeater%", "New one timer Sales%", "New Repeater Sales%","Old Repeater Sales%","New one timer Bills%","New Repeater Bills%","Old Repeater Bills%","Total Sales","Transaction Points Issued","Points Redeemed"]

    target_per = ["New one timer%", "New Repeater%", "Old Repeater%", "New one timer Sales%", "New Repeater Sales%","Old Repeater Sales%","New one timer Bills%","New Repeater Bills%","Old Repeater Bills%"]
    target_kpis_inr_cr = ["Total Sales"]
    target_kpis_lac = ["Transaction Points Issued","Points Redeemed"]

    target_columns = [df.columns[1], df.columns[2],df.columns[3]]

    for col in target_columns:
        df.loc[df["KPIS"].isin(target_per), col] = df.loc[df["KPIS"].isin(target_per), col].apply(lambda x: f"{round(x, 1)}%")

    for col in target_columns:
        df.loc[~df["KPIS"].isin(exceptions), col] = df.loc[~df["KPIS"].isin(exceptions), col].apply(format_with_locale)

    for col in target_columns:
        df.loc[df["KPIS"].isin(target_kpis_inr_cr), col] = df.loc[df["KPIS"].isin(target_kpis_inr_cr), col].apply(
            lambda x: f"{round(x / 10000000, 1)} Cr"
        )
    for col in target_columns:
        df.loc[df["KPIS"].isin(target_kpis_lac), col] = df.loc[df["KPIS"].isin(target_kpis_lac), col].apply(
            lambda x: f"{round(x / 100000, 1)} L"
        )
        
    desired_order = [
        "Enrollments", "Transacting Customers", "New one timer%", "New Repeater%", "Old Repeater%", "Total Sales","New one timer Sales%","New Repeater Sales%",
        "Old Repeater Sales%", "Total Bills","New one timer Bills%","New Repeater Bills%", "Old Repeater Bills%","ABV(Overall)","New one timer ABV","New repeater ABV",
        "Old repeater ABV","AMV(Overall)", "New one timer AMV","New repeater AMV","Old repeater AMV","Transaction Points Issued","Points Redeemed"
        ]
    
    df = df.set_index("KPIS") 
    df = df.loc[desired_order].reset_index()
    overall_abv = df.loc[df.iloc[:, 0] == "ABV(Overall)", df.columns[3]].values[0]
    return df,overall_abv







