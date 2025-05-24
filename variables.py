import calendar

# excel path
ppt_xl_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/Chogori_MBR_Deck/Data/chogori_mbr_mar_25_1.xlsx"
template_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/Chogori_MBR_Deck/Data/Chogori_MBR_template.pptx"
deck_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/Chogori_MBR_Deck/Data/chogori_mbr_MAR'25.pptx"
graph_path = "/home/ankit.sengar@corp.easyrewardz.com/Documents/Automation_deck/Chogori_MBR_Deck/Data/graph.png"

# current month number for which Deck needs to be created
curr_month_num = 3
# current year
curr_year = 2025
prev_month_num = (curr_month_num-1) % 12 or 12
prev_second_month_num = (curr_month_num-2) % 12 or 12
prev_third_month_num = (curr_month_num-3) % 12 or 12
prev_fourth_month_num = (curr_month_num-4) % 12 or 12

#twelve_months_back_num = (curr_month_num - 11) % 12 or 12
fy_year_start_month = 4

# previous year
prev_year = curr_year-1

# name assignment
curr_month_name = calendar.month_name[curr_month_num]
current_month = f"{curr_month_name.upper()[:3]}'{curr_year % 100}"

#current month last day
last_day = calendar.monthrange(curr_year, curr_month_num)[1]
if 10 <= last_day % 100 <= 20:
    suffix = "th"
else:
    suffix = {1: "st", 2: "nd", 3: "rd"}.get(last_day % 10, "th")

last_date = f"{last_day}{suffix} {curr_month_name.upper()[:3]}'{curr_year % 100}"

# name assignment
curr_month_prev_yr_name = calendar.month_name[curr_month_num]
curr_month_prev_yr = f"{curr_month_prev_yr_name.upper()[:3]}'{prev_year % 100}"

# month full name
curr_month_fname = calendar.month_name[curr_month_num]
curr_month_fname = f"{curr_month_fname}'{curr_year % 100}"

if curr_month_num==1:

    prev_month_name = calendar.month_name[prev_month_num]
    prev_month = f"{prev_month_name.upper()[:3]}'{prev_year % 100}"

    prev_second_month_name = calendar.month_name[prev_second_month_num]
    prev_second_month = f"{prev_second_month_name.upper()[:3]}'{prev_year % 100}"

    prev_third_month_name = calendar.month_name[prev_third_month_num]
    prev_third_month = f"{prev_third_month_name.upper()[:3]}'{prev_year % 100}"

    prev_fourth_month_name = calendar.month_name[prev_fourth_month_num]
    prev_fourth_month = f"{prev_fourth_month_name.upper()[:3]}'{prev_year % 100}"

elif curr_month_num==2:

    prev_month_name = calendar.month_name[prev_month_num]
    prev_month = f"{prev_month_name.upper()[:3]}'{curr_year % 100}"

    prev_second_month_name = calendar.month_name[prev_second_month_num]
    prev_second_month = f"{prev_second_month_name.upper()[:3]}'{prev_year % 100}"

    prev_third_month_name = calendar.month_name[prev_third_month_num]
    prev_third_month = f"{prev_third_month_name.upper()[:3]}'{prev_year % 100}"

    prev_fourth_month_name = calendar.month_name[prev_fourth_month_num]
    prev_fourth_month = f"{prev_fourth_month_name.upper()[:3]}'{prev_year % 100}"


elif curr_month_num==3:

    prev_month_name = calendar.month_name[prev_month_num]
    prev_month = f"{prev_month_name.upper()[:3]}'{curr_year % 100}"

    prev_second_month_name = calendar.month_name[prev_second_month_num]
    prev_second_month = f"{prev_second_month_name.upper()[:3]}'{curr_year % 100}"

    prev_third_month_name = calendar.month_name[prev_third_month_num]
    prev_third_month = f"{prev_third_month_name.upper()[:3]}'{prev_year % 100}"

    prev_fourth_month_name = calendar.month_name[prev_fourth_month_num]
    prev_fourth_month = f"{prev_fourth_month_name.upper()[:3]}'{prev_year % 100}"

elif curr_month_num==4:

    prev_month_name = calendar.month_name[prev_month_num]
    prev_month = f"{prev_month_name.upper()[:3]}'{curr_year % 100}"

    prev_second_month_name = calendar.month_name[prev_second_month_num]
    prev_second_month = f"{prev_second_month_name.upper()[:3]}'{curr_year % 100}"

    prev_third_month_name = calendar.month_name[prev_third_month_num]
    prev_third_month = f"{prev_third_month_name.upper()[:3]}'{curr_year % 100}"

    prev_fourth_month_name = calendar.month_name[prev_fourth_month_num]
    prev_fourth_month = f"{prev_fourth_month_name.upper()[:3]}'{prev_year % 100}"

else:
    
    prev_month_name = calendar.month_name[prev_month_num]
    prev_month = f"{prev_month_name.upper()[:3]}'{curr_year % 100}"

    prev_second_month_name = calendar.month_name[prev_second_month_num]
    prev_second_month = f"{prev_second_month_name.upper()[:3]}'{curr_year % 100}"

    prev_third_month_name = calendar.month_name[prev_third_month_num]
    prev_third_month = f"{prev_third_month_name.upper()[:3]}'{curr_year % 100}"

    prev_fourth_month_name = calendar.month_name[prev_fourth_month_num]
    prev_fourth_month = f"{prev_fourth_month_name.upper()[:3]}'{curr_year % 100}"

# finacial year string
if curr_month_num in (1, 2, 3):
    #ytd_curr_year_name = calendar.month_name[fy_year_start_month]
    #ytd_current_year = f"{ytd_curr_year_name[:3]}'{prev_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    fy = f"{prev_year % 100}-{curr_year % 100}"

else:
   # ytd_curr_year_name = calendar.month_name[fy_year_start_month]
   # ytd_current_year = f"{ytd_curr_year_name[:3]}'{curr_year % 100}-{curr_month_name[:3]}'{curr_year % 100}"
    fy = f"{curr_year % 100}-{curr_year % 100}"
