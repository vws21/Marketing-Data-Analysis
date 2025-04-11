import math
import pandas as pd

def task1(excel_file, sheet_name_placeholder):
    """
    TASK 1:
      1) Read file
      2) Extract columns 2 and 5
      3) For each sub-list:
         - Finds the '2/3 point' count value
         - Finds the 'first entry' count value
      4) Sum across all sub-lists:
         - sum_of_two_thirds (sum of 2/3-point entry counts)
         - sum_of_first (sum of first-entry counts)
      5) Compute percentage = (sum_of_two_thirds / sum_of_first) * 100
    """

    df_raw = pd.read_excel(
        excel_file,
        sheet_name=sheet_name_placeholder,
        header=0,
        skiprows=1,
        nrows=67
    )

    df_sub = df_raw.iloc[:, [1, 4]].copy()
    df_sub.columns = ['sub_list_id', 'count']

    grouped = df_sub.groupby('sub_list_id')

    sum_of_two_thirds = 0
    sum_of_first = 0
    two_thirds_dict = {}

    for sub_list_id, group_data in grouped:
        n = len(group_data)
        idx_1b = math.ceil((2/3) * n)
        idx_0b = idx_1b - 1

        # "2/3" sub-count
        two_thirds_count = group_data.iloc[idx_0b]['count']
        two_thirds_dict[sub_list_id] = two_thirds_count
        sum_of_two_thirds += two_thirds_count

        # "First" sub-count
        first_count = group_data.iloc[0]['count']
        sum_of_first += first_count

    percentage_2_3 = 0
    if sum_of_first != 0:
        percentage_2_3 = (sum_of_two_thirds / sum_of_first) * 100

    return {
        "two_thirds_dict_per_sublist": two_thirds_dict,
        "sum_of_two_thirds": sum_of_two_thirds,
        "sum_of_first": sum_of_first,
        "percentage_of_two_thirds_over_first": percentage_2_3
    }

def task2(excel_file, sheet_name_placeholder):
    """
    TASK 2:
      1) Columns: 2nd = sub_list_id, 5th = sub_count
      2) Find fraction of 'sum_of_first' over 'accepted'
      3) Find '1/2 point' aggregate sub-count and compares to sum_of_first
      4) Calculate final 'percent of a percent'
    """

    df_raw = pd.read_excel(
        excel_file,
        sheet_name=sheet_name_placeholder,
        header=0,
        skiprows=1,
        nrows=67
    )

    df_sub = df_raw.iloc[:, [1, 4]].copy()
    df_sub.columns = ['sub_list_id', 'count']

    grouped = df_sub.groupby('sub_list_id')
    sum_of_first = 0
    for _, group_data in grouped:
        first_count = group_data.iloc[0]['count']
        sum_of_first += first_count

    accepted_count = 3017 # Count derived from "Prospect Data - 2023 to 2025" sheet
    if sum_of_first != 0:
        accepted_fraction = accepted_count / sum_of_first
    accepted_percentage = accepted_fraction * 100

    sum_of_half = 0
    grouped = df_sub.groupby('sub_list_id')
    for _, group_data in grouped:
        n = len(group_data)
        idx_1b_half = math.ceil(0.5 * n)
        idx_0b_half = idx_1b_half - 1
        half_count = group_data.iloc[idx_0b_half]['count']
        sum_of_half += half_count

    half_percentage = 0
    if sum_of_first != 0:
        half_percentage = (sum_of_half / sum_of_first) * 100

    final_percent_of_percent = (accepted_percentage * half_percentage) / 100

    return {
        "sum_of_first": sum_of_first,
        "accepted_count": accepted_count,
        "accepted_fraction": accepted_fraction,
        "accepted_percentage": accepted_percentage,
        "sum_of_half": sum_of_half,
        "half_percentage_of_first": half_percentage,
        "final_percent_of_percent": final_percent_of_percent
    }

def task3(excel_file, sheet_name_placeholder):
    """
    TASK 3:
        1. Identical in execution to task 1
    """

    df_raw = pd.read_excel(
        excel_file,
        sheet_name=sheet_name_placeholder,
        header=0,
        skiprows=1,
        nrows=91
    )

    df_sub = df_raw.iloc[:, [1, 4]].copy()
    df_sub.columns = ['sub_list_id', 'count']

    grouped = df_sub.groupby('sub_list_id')

    sum_of_two_thirds = 0
    sum_of_first = 0
    two_thirds_dict = {}

    for sub_list_id, group_data in grouped:
        n = len(group_data)
        idx_1b = math.ceil((2/3) * n)
        idx_0b = idx_1b - 1

        # "2/3" sub-count
        two_thirds_count = group_data.iloc[idx_0b]['count']
        two_thirds_dict[sub_list_id] = two_thirds_count
        sum_of_two_thirds += two_thirds_count

        # "First" sub-count
        first_count = group_data.iloc[0]['count']
        sum_of_first += first_count

    percentage_2_3 = 0
    if sum_of_first != 0:
        percentage_2_3 = (sum_of_two_thirds / sum_of_first) * 100

    return {
        "two_thirds_dict_per_sublist": two_thirds_dict,
        "sum_of_two_thirds": sum_of_two_thirds,
        "sum_of_first": sum_of_first,
        "percentage_of_two_thirds_over_first": percentage_2_3
    }

if __name__ == "__main__":

    excel_file = "relevant.xlsx"
    
    # Task 1:
    sheet_task1 = "Prospect Journey BY EMAIL"
    results_task1 = task1(excel_file, sheet_task1)
    print("\nTASK 1:")
    print("---------------")
    print(f"\nSum of 2/3 counts = {results_task1['sum_of_two_thirds']}")
    print(f"Sum of first-entry counts = {results_task1['sum_of_first']}")
    print(f"Percentage (2/3 sum over first sum) = {results_task1['percentage_of_two_thirds_over_first']:.2f}%")

    # Task 2:
    results_task2 = task2(excel_file, sheet_task1)
    print("\nTASK 2:")
    print("---------------")
    print(f"\nAccepted recipients = {results_task2['accepted_count']}")
    print(f"Accepted fraction (accepted / sum_of_first) = {results_task2['accepted_fraction']:.4f}")
    print(f"Accepted percentage = {results_task2['accepted_percentage']:.2f}%")
    print(f"Sum of 1/2-entry counts = {results_task2['sum_of_half']}")
    print(f"Percentage of 1/2-entry sum over first sum = {results_task2['half_percentage_of_first']:.2f}%")
    print(f"Final percent = {results_task2['final_percent_of_percent']:.2f}%")

    # Task 3:
    sheet_task3 = "Admit Journey BY EMAIL"
    results_task3 = task3(excel_file, sheet_task3)
    print("\nTASK 3:")
    print("---------------")
    print(f"\nSum of 2/3 counts = {results_task3['sum_of_two_thirds']}")
    print(f"Sum of first-entry counts = {results_task3['sum_of_first']}")
    print(f"Percentage (2/3 sum over first sum) = {results_task3['percentage_of_two_thirds_over_first']:.2f}%")
