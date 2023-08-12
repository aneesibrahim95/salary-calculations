def split(time) :
    hour = ""
    minute = ""

    for i in time:
        try:
            if len(hour) < 2:
                b = int(i)
                hour += i
            elif len(hour) >= 2:
                if len(minute) < 2:
                    b = int(i)
                    minute += i
        except ValueError:
            pass
    hour = int(hour)
    minute = int(minute)
    return hour, minute




# This funciton takes in the name of a person and returns the
#  - Total work duration (In Hours)
#  - Total working days (Both Punch in and Punch out should be there)
#  - Punch in Punch out details (Time of in and out)
def workdur(pandas_library, dataframe, name):
    target_value = name

    locations = []
    # Loop through each row and column to scan all cells
    for row_index, row in dataframe.iterrows():
        for col_index in dataframe.columns:
            cell_value = row[col_index]
            try:
                if target_value.lower() in str(cell_value.lower()):
                    locations.append(row_index)
                    locations.append(col_index)
            #         now, the locations list will contain the [row number, column number] of the searched person's cell
            except AttributeError:
                pass
    # print (name)

    employee_name = dataframe.loc[locations[0],locations[1]]
    # (from the cell location data,the name of the employee is obtained from the biometric sheet)
    print (employee_name)

    month_length = 300 #A random number
    intime_dict = {}
    date = []
    date_counter = 1

    for i in range(month_length):
        try:
            status_row = ((locations[0]) + 1)                       # Row number of the status row : Fixed
            status_col = dataframe.columns[i]                       # Column number of the status row : Being iterated
            status_start = [(status_row),(status_col)]
            status = dataframe.loc[status_start[0],status_start[1]] # Status obtained
            # print(status)
            if status == "A" or status == "P":
                date.append(date_counter)
                date_counter += 1
                intime_row = ((locations[0]) + 2)
                intime_col = dataframe.columns[i]
                intime_start = [(intime_row), (intime_col)]

                # if pandas_library.notna(dataframe.loc[intime_start[0], intime_start[1]]):
                intime = dataframe.loc[intime_start[0], intime_start[1]]
                # else:
                #     pass
                # if pandas_library.notna(dataframe.loc[intime_start[0], intime_start[1]]):
                intime_dict[i + 1] = intime
                # else:
                #     pass

                # print(intime)
            else :
                month_length += 1
        except KeyError :
            pass
        except SyntaxError :
            pass
        except IndexError :
            pass
        finally:
            pass
    #
    outtime_dict = {}

    for i in range(month_length):
        try:
            status_row = ((locations[0]) + 1)
            status_col = dataframe.columns[i]
            status_start = [(status_row),(status_col)]
            status = dataframe.loc[status_start[0],status_start[1]]
            if status == "A" or "P" :
                outtime_row = ((locations[0]) + 3)
                outtime_col = dataframe.columns[i]
                outtime_start = [(outtime_row), (outtime_col)]

                # if pandas_library.notna(dataframe.loc[outtime_start[0], outtime_start[1]]):
                outtime = dataframe.loc[outtime_start[0], outtime_start[1]]
                # else:
                #     pass
                # if pandas_library.notna(dataframe.loc[outtime_start[0], outtime_start[1]]):
                outtime_dict[i + 1] = outtime
                # else:
                #     pass
            else :
                month_length += 1
        except KeyError :
            pass
        except SyntaxError :
            pass
        except IndexError:
            pass
        finally:
            pass
    #
    print(intime_dict)
    print(len(intime_dict))
    print(outtime_dict)
    print(len(outtime_dict))

    tw_dur = 0
    det_work_dur = []
    counter = 0
    date = []
    zero_counter = 0

    for i in range(300):
        try:
            status_row = ((locations[0]) + 1)
            status_col = dataframe.columns[i]
            status_start = [(status_row), (status_col)]
            status = dataframe.loc[status_start[0], status_start[1]]
            # print(status)
            # if (pandas_library.notna(intime_dict[(i+1)]) and pandas_library.notna(outtime_dict[(i+1)])): # Checking if the intime and outtime are there, avoiding those with no values
            a = intime_dict[(i+1)]                                     # Splitting intime and outtime Hours and Minutes
            b = outtime_dict[(i+1)]

            if pandas_library.notna(a):
                in_hour, in_minute = split(a)
            else:
                in_hour = 0
                in_minute = 0

            if pandas_library.notna(b) :
                out_hour, out_minute = split(b)
            else:
                out_hour = 0
                out_minute = 0

            # print (f" {i}outtime : {bb}")

            if (in_hour==0 and in_minute==0) or (out_hour==0 and out_minute==0):
                c = 0
                zero_counter += 1

            else:
                k = (out_hour - in_hour)*60 + (out_minute - in_minute)
                c = k/60
            print(in_hour, in_minute, out_hour, out_minute, c)
            # Total work duration of every day is calculated

            if status == "A" or status == "P":
                counter += 1

            det_work_dur.append([counter, a, b , (round(c,2))])                           #

            tw_dur = tw_dur + c
            # tw_days += 1
        except KeyError:
            # print (f"key error {i}")
            pass
        except IndexError:
            # print(f"Index error {i}")
            pass



    tw_days = min(len(intime_dict), len(outtime_dict)) - zero_counter
    print(det_work_dur)
    print(f"{name} was {tw_dur} hours working")
    print(f"{name} was present on {tw_days} days as per minumum of punch in details")
    print(f"number of days present as per detailed work report : {len(det_work_dur)}")
    return employee_name,tw_dur,tw_days,det_work_dur


def add_table_to_doc(document, employee_name, output_structure):
    # Add a heading for the employee name
    document.add_heading(employee_name, level=1)

    # Create a table with three columns for "Parameter", "Value", and "Pay"
    table = document.add_table(rows=1, cols=3)
    table.autofit = True

    # Add data to the table
    for key, value in output_structure.items():
        row_cells = table.add_row().cells
        row_cells[0].text = key  # Display the key (parameter)
        row_cells[1].text = str(value[0])  # Convert value to string
        row_cells[2].text = str(value[1])  # Convert pay to string

