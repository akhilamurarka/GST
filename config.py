# timeout for finding the file name
TIMEOUT=30

# Month and quarter mapping
month_name_to_number = {
    "January": 1, "Jan": 1, "February": 2, "Feb": 2, "March": 3, "Mar": 3,
    "April": 4, "Apr": 4, "May": 5, "June": 6, "Jun": 6,
    "July": 7, "Jul": 7, "August": 8, "Aug": 8, "September": 9, "Sep": 9,
    "October": 10, "Oct": 10, "November": 11, "Nov": 11, "December": 12, "Dec": 12,
}
quarter_text_map = {
    "1": "Quarter 1 (Apr - Jun)", "2": "Quarter 2 (Jul - Sep)",
    "3": "Quarter 3 (Oct - Dec)", "4": "Quarter 4 (Jan - Mar)"
}

month_small_to_big={
'Jan':'January',
'Feb':'February',
'Mar':'March',
'Apr':'April',
'May':'May',
'Jun':'June',
'Jul':'July',
'Aug':'August',
'Sep':'September',
'Oct':'October',
'Nov':'November',
'Dec':'December'
}    

def get_quarter_from_month(month_num):
    if 1 <= month_num <= 3:
        return "4"
    elif 4 <= month_num <= 6:
        return "1"
    elif 7 <= month_num <= 9:
        return "2"
    elif 10 <= month_num <= 12:
        return "3"
    return None