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
                
                        if (row<=1) or (row>=0 and col==0):
                            pass
                        else:
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

                            if row > 1 and col>0: # leaving headers value
                                new_run.text = str(df.iloc[row-2,col-1])

                        
    return slide

def create_df(chogori_ppt_xl_path,sheet_name):
    df = pd.read_excel(chogori_ppt_xl_path,sheet_name = sheet_name)
    df_male = df[["male_item_code","male_item_name","male_units_sold"]]
    df_male = df_male.sort_values(by="male_units_sold", ascending=False).reset_index(drop=True)
    df_male = df_male.head(10)

    df_female = df[["female_item_code","female_item_name","female_units_sold"]]
    df_female = df_female.sort_values(by="female_units_sold", ascending=False).reset_index(drop=True)
    df_female = df_female.head(10)

    df = pd.concat([df_male,df_female],axis = 1)
    df[df.columns[2]] = df[df.columns[2]].apply(format_with_locale)
    df[df.columns[5]] = df[df.columns[5]].apply(format_with_locale)
    return df







