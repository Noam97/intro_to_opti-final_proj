import pandas
import pulp
import numpy as np

NUM_RESIDENTS = 25
MONTH_DAYS = 30
DATES = ['Jan-01','Jan-02','Jan-03','Jan-04','Jan-05','Jan-06','Jan-07','Jan-08','Jan-09','Jan-10','Jan-11','Jan-12',
         'Jan-13','Jan-14','Jan-15','Jan-16','Jan-17','Jan-18','Jan-19','Jan-20','Jan-21','Jan-22','Jan-23','Jan-24',
         'Jan-25','Jan-26','Jan-27','Jan-28','Jan-29','Jan-30']

def create_obj_func(data, prob):
    id = 0
    # add variables
    for resident in data.keys():
        data[resident]["my_hospital_shifts"] = []
        data[resident]["my_clinics"] = []
        for day in range(MONTH_DAYS):
            data[resident]["my_hospital_shifts"].append(pulp.LpVariable("x_{}_{}".format(str(id), str(day)), cat=pulp.LpBinary))
            data[resident]["my_clinics"].append(pulp.LpVariable("y_{}_{}".format(str(id), str(day)), cat=pulp.LpBinary,upBound=binary_clinics[day]))
        id += 1
    # create objective function which it's goal is to minimize the uncomfort shifts for resident for whole month
    objective_func = None
    for resident in data.keys():
        obj = None
        for x in range(len(data[resident]["my_hospital_shifts"])):
            obj += data[resident]["preferences"][x]*data[resident]["my_hospital_shifts"][x]
        objective_func += obj
    prob += objective_func
    return prob, data, clinics, num_of_clinics

def add_hospital_constraints(prob, data):
    # Every day we need exactly 3 residents for hospital shifts
    for day in range(MONTH_DAYS):
        constraint = None
        for resident in data.keys():
            constraint += data[resident]["my_hospital_shifts"][day]
        prob += constraint == 3

    # Sum of all residents degree that work on specific day has to be at least 6
    for day in range(MONTH_DAYS):
        constraint = None
        for resident in data.keys():
            constraint += data[resident]["my_hospital_shifts"][day] * data[resident]["degree"]
        prob += constraint >= 6

    # Resident can work at most 6 hospital shifts in a month
    for resident in data.keys():
        constraint = None
        for day in range(MONTH_DAYS):
            constraint += data[resident]["my_hospital_shifts"][day]
        prob += constraint <= 6

    # a resident can't work day on day off. At least 2 shifts free between residents shift
    for resident in data.keys():
        for day in range(MONTH_DAYS - 2):
            prob += data[resident]["my_hospital_shifts"][day] + \
                    data[resident]["my_hospital_shifts"][day + 1] + data[resident]["my_hospital_shifts"][day + 2] <= 1

    # The amount of days a resident has hospital shifts in a month will be +-average
    average = 90 / NUM_RESIDENTS
    for resident in data.keys():
        constraint = None
        for day in range(MONTH_DAYS):
            constraint += data[resident]["my_hospital_shifts"][day]
        prob += constraint <= average + 1
        prob += constraint >= average - 1
    return prob

def add_clinic_constraints(prob, data, clinics, num_of_clinics):
    # When filling a clinic shift, resident cannot have a hospital shift on same day or day before
    for resident in data.keys():
        for day in range(MONTH_DAYS):
            if day==0:
                 prob += data[resident]["my_clinics"][day]+data[resident]["my_hospital_shifts"][day]<=1
            else:
                prob += data[resident]["my_clinics"][day] + 0.5*data[resident]["my_hospital_shifts"][day] + \
                           0.5*data[resident]["my_hospital_shifts"][day - 1] <= 1

    # number of residents that do a clinic shift must be exactly the number of residents needed at that day in clinics
    for day in range(MONTH_DAYS):
        sum_residents_per_day = None
        for resident in data.keys():
            sum_residents_per_day += data[resident]["my_clinics"][day]
        prob += sum_residents_per_day == clinics[day]

    # The amount of days a resident has clinic shifts in a month will be +-average
    average_clinics = num_of_clinics / NUM_RESIDENTS
    for resident in data.keys():
        constraint = None
        for day in range(MONTH_DAYS):
            constraint += data[resident]["my_clinics"][day]
        prob += constraint <= average_clinics + 1
        prob += constraint >= average_clinics - 1
    return prob


if __name__ == "__main__":
    # read excels
    clinics_excel = pandas.read_excel("clinics.xlsx", header=0)
    resident_preferences = pandas.read_excel("residents.xlsx", header=0)
    # order clinics data from excel
    num_of_clinics = 0 # to count total number of clinic shifts for whole month
    clinics = None
    binary_clinics = []
    for row in clinics_excel.iterrows():
        clinics = row[1]
    # set a binary list for dates with clinic shifts requirements
    for index in range(len(clinics)):
        if clinics[index] > 0:
            binary_clinics.append(1)
        else:
            binary_clinics.append(0)
        num_of_clinics += int(clinics[index])
    # order all data
    data = {}
    for row in resident_preferences.iterrows():
        resident_id = row[1][0]
        data[resident_id] = {}
        data[resident_id]["degree"] = row[1][1]
        data[resident_id]["preferences"] = []
        for day in range(MONTH_DAYS):
            data[resident_id]["preferences"].append(int(row[1][day+2]))
    prob = pulp.LpProblem("ResidentsShifts", pulp.LpMinimize)
    # create objective function
    prob, residents_data, clinics, num_of_clinics = create_obj_func(data, prob)
    # add hospital constraints
    prob = add_hospital_constraints(prob, residents_data)
    # add clinics constraints
    prob = add_clinic_constraints(prob, residents_data, clinics, num_of_clinics)

    try:
        prob.solve()
    except Exception as e:
        print("Couldn't find a solution: {}".format(e))

    month_scheduling = []
    clinic_scheduling = []
    for day in range(MONTH_DAYS):
        month_scheduling.append([])
        clinic_scheduling.append([])
    for resident in residents_data.keys():
        for d in range(MONTH_DAYS):
            if residents_data[resident]["my_hospital_shifts"][d].varValue == 1:
                month_scheduling[d].append(resident)
            if residents_data[resident]["my_clinics"][d].varValue == 1:
                clinic_scheduling[d].append(resident)

    x = 1
    # order residents of each day in month by their degree
    for day in month_scheduling:
        degrees = []
        for d in day:
            degrees.append(residents_data[d]["degree"])
        npArr = np.array(degrees)
        sorted_indexes = np.argsort(npArr)
        month_scheduling[x - 1] = [day[sorted_indexes[0]], day[sorted_indexes[1]], day[sorted_indexes[2]]]
        x += 1

    i=1
    clinic = []
    # order residents of each day in month by their degree
    for x in clinic_scheduling:
        clinic.append([])
        if len(x)!=0 :
            for j in range(len(x)):
                clinic[i-1].append(x[j])
        i+=1

    # write to excel sheet
    for k in range(len(month_scheduling)):
        if clinic[k] == []:
            month_scheduling[k].append("")
        else:
            month_scheduling[k].append(clinic[k])
    shift =['Yoldot', 'Nashim', 'Leda', 'clinics']
    df1 = pandas.DataFrame(month_scheduling,
                       index=DATES,
                       columns=shift)
    df1.to_excel("ResidentMonthShifts.xlsx")