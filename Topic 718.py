import decimal
import math
import csv
import xlsxwriter


def summation(r, v, c, l, g):
    # This function calculates the "Expected Term"
    # This function does not calculate cliff situation.
    # If there is a cliff that is handled by "Stage1"
    a = 0
    i = (c/r)+1
    while i <= v/r:
        a += i*r*((1-g)/(v/r))
        i += 1
    a += l
    return a


def band_assign():
    # changing this comment so I can test github
    if exp_term >= 7:
        band = 5
        return band
    elif exp_term >= 5:
        band = 4
        return band
    elif exp_term >= 3:
        band = 3
        return band
    elif exp_term >= 2:
        band = 2
        return band
    elif exp_term >= 1:
        band = 1
        return band


def calc_rate(band, one_rate, two_rate, three_rate, five_rate, seven_rate, ten_rate, exp_term):
    rate = 0
    slope = 0
    x = 0
    y = 0
    if band == 1:
        slope = float(1/((2 - 1)/(two_rate - one_rate)))
        x = 2
        y = two_rate
    elif band == 2:
        slope = float(1/((3 - 2)/(three_rate - two_rate)))
        x = 3
        y = three_rate
    elif band == 3:
        slope = float(1/((5 - 3)/(five_rate - three_rate)))
        x = 5
        y = five_rate
    elif band == 4:
        slope = float(1/((7 - 5)/(seven_rate - five_rate)))
        x = 7
        y = seven_rate
    elif band == 5:
        slope = float(1/((10 - 7)/(ten_rate - seven_rate)))
        x = 10
        y = ten_rate
    rate = (slope*exp_term) + (y - (slope*x))
    return rate


def d1_calc(fmv, exp_price, exp_term, rate, div, vol):
    stage1 = math.log(fmv/exp_price)
    stage2 = (rate - div + (vol * vol * 0.5)) * exp_term
    stage3 = vol * math.sqrt(exp_term)
    d1 = stage1+stage2/stage3
    return d1


def d2_calc(d1, vol, exp_term):
    d2 = d1 - (vol * math.sqrt(exp_term))
    return d2


def black_calc():
    stage4 = fmv * math.exp(-div * exp_term)
    stage5 = (1.0 + math.erf(d1 / math.sqrt(2.0))) / 2.0
    stage6 = exp_price * math.exp(-rate * exp_term)
    stage7 = (1.0 + math.erf(d2 / math.sqrt(2.0))) / 2.0
    c = (stage4*stage5)-(stage6*stage7)
    return c


def calc_compound():
    m = 1-(1 - y)**(1/12)
    return m


def summation2(r, v, c, g, tg, m):
    a = float(0.00)
    d = float(0.00)
    i = (c/r)+1
    while i <= v/r:
        x = (tg-(tg*g))/(v/r)
        z = ((1 - m)**(i*r))
        d += (x * z)
        # print(((v-c)/r)+c)
        # print(x)
        # print(z)
        # print(d)
        # print(i)
        # print("BREAK")
        i += 1
    return d


def amortize(fair_val, v, ac):
    # Prints out how much you should expense for each month of the vesting schedule
    i = decimal.Decimal(0)
    a = decimal.Decimal(0)
    v = decimal.Decimal(v)
    fair_val = decimal.Decimal(fair_val)
    ac = decimal.Decimal(ac)

    while i < v:
        # Straight Line Method
        if (fair_val - a) * (1 / (v - i)) > (fair_val - a) * (1 / v) * ac:
            print("Month {0} Expense ${1:.2f}".format(i + 1, (fair_val - a) * (1 / (v - i))))
            a += (fair_val - a) * (1 / (v - i))
            i += 1
        else:
            # Modified Straight Line
            print("Month {0} Expense ${1:.2f}".format(i + 1, (fair_val - a) * (1 / v) * ac))
            i += 1
            a += (fair_val - a) * (1 / v) * ac


def print_stats():
    print("r = {0}, v = {1}, c = {2}, l = {3}, g = {4}".format(r,v,c,l,g))
    print("stage1 = {}".format(stage1))
    print("stage2 = {}".format(stage2))
    print("Expected Term (Monthly) = {}".format(mo_exp_term))
    print("Expected Term (Yearly) = {}".format(yr_exp_term))
    print("Your calculated interest rate is: {}%".format(rate))
    print("d1 = {}".format(d1))
    print("d2 = {}".format(d2))
    print("Calculated share value = ${}".format(black))
    print("stage1 = {}".format(stage4))
    print("stage2 = {}".format(stage5))
    print("Expected to Vest = {}".format(exp_vest))
    print("Actual Fair Value = {}".format(tg*black))
    print("Total Projected Fair Value = {}".format(fair_val))


def get_average(grant_num):
    i = 0
    grant_total = 0
    total_num_options = 0
    while i < grant_num:
        grant_total += float(black_volatility_list[i]) * float(total_options[i])
        total_num_options += float(total_options[i])
        i += 1
    return grant_total / total_num_options, total_num_options


def get_excel_period_dates():
    days = 0
    month = (period_end[4:])
    if month == "01":
        days = 31
    elif month == "02":
        days = 28
    elif month == "03":
        days = 31
    elif month == "04":
        days = 30
    elif month == "05":
        days = 31
    elif month == "06":
        days = 30
    elif month == "07":
        days = 31
    elif month == "08":
        days = 31
    elif month == "09":
        days = 30
    elif month == "10":
        days = 31
    elif month == "11":
        days = 30
    elif month == "12":
        days = 31
    excel_end_date = "{}/{}/{}".format(period_end[4:], days, period_end[0:5])
    excel_start_date = "{}/01/{}".format(period_start[4:], period_end[0:5])
    return excel_start_date, excel_end_date

# Collect the account period date range we are going to change some code and see if anything happens in github
period_start = input("What is the starting date of this reporting period? (yyyymm)> ")
period_end = input("What is the ending date of this reporting period? (yyyymm)> ")

# This function turns the yyyymm format into mm/dd/yyyy format. This is used in printing to the Excel file ddd
excel_start_date, excel_end_date = get_excel_period_dates()

# Opens the csv file
f = open('input.csv')
csv_f = csv.reader(f)

# Creating lists that will store scv values by category (column) testing testing
grant_start_month = []
grant_start_year = []
tranche_res = []
vest_length = []
cliff_length = []
option_life = []
pre_vested = []
fair_market = []
exe_price = []
dividend = []
volatility = []
total_options = []
annual_for = []
accelerator = []
year_1_int = []
year_2_int = []
year_3_int = []
year_5_int = []
year_7_int = []
year_10_int = []

# Lists that are used only for Valuation on Disclosure #1 (only include values applicable to reporting period)
total_int_rate = []
weighted_int_rate = []
exp_term_list = []
exp_vest_list = []
weighted_exp_term = []
weighted_dividend = []
fair_market_list = []
weighted_fair_market = []
total_outstanding_list = []
total_exercise_list = []
total_forfeitures_list = []
total_vested_list = []
black_volatility_list = []
black_dividend = []
period_total_grants = []
period_total_exercised = []
period_total_forfeited = []
period_expirations = []
period_outstanding = []
end_period_exercisable = []

# Lists that include values applicable BEFORE the reporting period. Prefix = "pre"
pre_total_outstanding = []
pre_total_exercised = []
pre_total_forfeited = []

# Initialize counters
i = 0
black_count = 0
weighted_outstanding = 0
weighted_grants = 0
weighted_exercised = 0
weighted_forfeitures = 0
weighted_expirations = 0
weighted_exercisable = 0
weighted_vested = 0
weighted_unvested = 0
weighted_outstanding_term = 0
weighted_exercisable_term = 0
weighted_unvested_term = 0
weighted_vested_term = 0
total_unvested = 0
contract_term = 0

# Compute all variables required for Black-Scholes model. Computed for every grant.
# Compute Black-Scholes fair market value. Computed for every grant.
for row in csv_f:
    # This captures data from the input spreadsheet and organizes it
    # into lists by column (category)
    grant_start_month.append(row[0][:2])
    grant_start_year.append(row[0][6:])
    total_outstanding_list.append(row[7])
    total_exercise_list.append(row[6])
    total_forfeitures_list.append(row[5])
    tranche_res.append(row[10])
    vest_length.append(row[11])
    cliff_length.append(row[13])
    option_life.append(row[14])
    pre_vested.append(row[12])
    fair_market.append(row[9])
    exe_price.append(row[8])
    dividend.append(row[15])
    volatility.append(row[16])
    total_options.append(row[4])
    annual_for.append(row[18])
    accelerator.append(row[17])
    year_1_int.append(row[19])
    year_2_int.append(row[20])
    year_3_int.append(row[21])
    year_5_int.append(row[22])
    year_7_int.append(row[23])
    year_10_int.append(row[24])

    # The computational piece of the code was written with the following
    # variables. This portion of code just maps those existing variables
    # to the applicable data from the input spreadsheet.
    r = float(tranche_res[i])
    v = float(vest_length[i])
    c = float(cliff_length[i])
    l = float(option_life[i])
    g = float(pre_vested[i])
    fmv = float(fair_market[i])
    exp_price = float(exe_price[i])
    div = float(dividend[i])
    vol = float(volatility[i])
    tg = float(total_options[i])
    y = float(annual_for[i])
    ac = float(accelerator[i])
    one_rate = float(year_1_int[i])
    two_rate = float(year_2_int[i])
    three_rate = float(year_3_int[i])
    five_rate = float(year_5_int[i])
    seven_rate = float(year_7_int[i])
    ten_rate = float(year_10_int[i])

    # This is the first portion of calculating the "Expected Term" (months)
    # This equation only provides a value other than zero when there is a cliff in the vesting terms.
    stage1 = c*((1-g)*(c/v))

    # This is the second portion of calculating the "Expected Term" (months)
    # This equation always provides a value other than zero assuming the grant is valid.
    stage2 = summation(r, v, c, l, g)
    decimal.getcontext().prec = 10

    # dividing by two is a practice sanctioned/mandated by the IRS.
    mo_exp_term = (stage1 + stage2)/2
    yr_exp_term = mo_exp_term/12
    exp_term = float(yr_exp_term)

    # This identifies the applicable interest rates based upon the previously calculated "Expected Term"
    band = band_assign()

    # Once we know which two interest rates to use as reference points we can continue
    # to solve for Y (the actual interest rate)
    rate = float(calc_rate(band, one_rate, two_rate, three_rate, five_rate, seven_rate, ten_rate, exp_term))

    # This function call calculates d1 of the Black-Scholes equation
    d1 = d1_calc(fmv, exp_price, exp_term, rate, div, vol)

    # This function call calculates d2 of the Black-Scholes equation
    d2 = d2_calc(d1, vol, exp_term)

    # This function uses d1 and d2 in the modified Black-Scholes equation
    # to determine Fair Market share value
    black = black_calc()

    # This function call returns the monthly compound interest to
    # reach the annual forfeiture rate in a year (dynamically).
    m = calc_compound()

    # "Stage4" is the first portion of calculating the number of shares "Expected to Vest"
    # This equation only provides a value other than zero when there is a cliff in the vesting terms.
    stage4 = ((1-m)**c)*(((tg-(tg*g))/v)*c)

    # "Stage5" is the second portion of calculating the number of shares "Expected to Vest"
    stage5 = summation2(r, v, c, g, tg, m)

    # In the event there are pre-vested shares "Stage6" is the third portion of calculating the number of shares
    # "Expected to Vest"
    stage6 = tg*g
    exp_vest = stage4 + stage5 + stage6

    # Calculates the total value of shares expected to vest per grant
    fair_val = exp_vest * black

    # These lists are used in calculating the valuation portion of disclosure #1
    # Essentially, they only accept information from grants made in the disclosure period
    grant_start_date = float("{}{}".format(grant_start_year[i], grant_start_month[i]))
    if grant_start_date >= float(period_start):
        if grant_start_date <= float(period_end):
            black_volatility_list.append(vol)
            total_int_rate.append(rate)
            weighted_int_rate.append(rate * tg)
            exp_term_list.append(exp_term)
            weighted_exp_term.append(exp_term * tg)
            black_dividend.append(div)
            weighted_dividend.append(div * tg)
            fair_market_list.append(black)
            weighted_fair_market.append(black * tg)
            period_total_grants.append(row[4])
            period_total_exercised.append(row[6])
            period_total_forfeited.append(row[5])
            period_outstanding.append(row[7])
            weighted_grants += float(row[4]) * float(row[8])
            weighted_exercised += float(row[6]) * float(row[8])
            weighted_forfeitures += float(row[5]) * float(row[8])
            black_count += 1
    else:
        # These are lists that contain variables with grant dates BEFORE the reporting period
        pre_total_outstanding.append(row[4])
        pre_total_exercised.append(row[6])
        pre_total_forfeited.append(row[5])

        # weighted outstanding shares = (shares granted - (exercised + forfeited)) * exercise price
        weighted_outstanding += (float(row[4]) - (float(row[6]) + float(row[5]))) * float(row[8])

    # converting csv grant date into a numerical value for comparison. Format = yyyymm
    grant_expiration_date = float("{}{}".format(row[1][6:], row[1][:2]))

    # This set of "if" statements determines whether or not an expiration date occurred during a reporting period
    if grant_expiration_date >= float(period_start):
        if grant_expiration_date <= float(period_end):
            period_expirations.append(total_options)
            weighted_expirations = (float(row[4]) - float(row[6])) * float(row[8])

    vested = 0
    raw_term_difference = 0
    term_difference = 0
    # This "if" statement determines how long a particular grant has been vesting (minus cliff, effective vesting)
    if grant_expiration_date >= float(period_end) and row[5] == "0":
        grant_start_date_
        s = (float(row[0][6:]) * 12) + float(row[0][:2])
        period_end_date_months = (float(period_end[0:4]) * 12) + float(period_end[4:])
        vesting_period_months = (period_end_date_months - grant_start_date_months) - 1
        print("Number of months that vesting has been able to accrue {}".format(vesting_period_months))
        vested = (tg * g)
        date_counter = r

        # This "while" loop adds up the number of shares vested to date for any ONE grant that enters this loop.
        while date_counter < vesting_period_months - (c - 1):
            vested += (tg - (tg * g)) / (v - c)
            date_counter += r

    if grant_expiration_date > float(period_end) and row[5] == "0":
        # This is the total number of vested shares minus the shares that have already been exercised.
        end_period_exercisable.append(vested - float("{}".format(row[6])))
        weighted_exercisable += (vested - float("{}".format(row[6]))) * float(row[8])
        weighted_vested += vested * float(row[8])
        weighted_unvested += (float(row[7]) - vested) * float(row[8])
        total_unvested += float(row[7]) - vested
        raw_term_difference = grant_expiration_date - float(period_end)



        contract_term += vested * term_difference
        total_vested_list.append(vested)
        print("ATTENTION contract_term = {}".format(contract_term))
    # weighted_vested_term += vested *

    # Prints crucial information to console
    print_stats()

    # Prints the amortization schedule per grant
    amortize(fair_val, v, ac)

    i += 1
    k = 0

# calculate the values required by the valuation section of disclosure #1
vol_low = float(min(black_volatility_list))
vol_high = float(max(black_volatility_list))
vol_average, total_num_options = get_average(black_count)

int_rate_low = float(min(total_int_rate))
int_rate_high = float(max(total_int_rate))
int_rate_average = float(sum(weighted_int_rate) / total_num_options)

exp_term_low = float(min(exp_term_list))
exp_term_high = float(max(exp_term_list))
exp_term_average = float(sum(weighted_exp_term) / total_num_options)

dividend_low = float(min(black_dividend))
dividend_high = float(max(black_dividend))
dividend_total = float(sum(weighted_dividend))
dividend_average = (dividend_total / total_num_options)

fair_market_low = float(min(fair_market_list))
fair_market_high = float(max(fair_market_list))
fair_market_total = float(sum(weighted_fair_market))
fair_market_average = (fair_market_total / total_num_options)


print("")
print("Volatility Range = {}% - {}%".format(vol_low * 100, vol_high * 100))
print("Volatility Weighted Average = {:.2f}%".format(vol_average * 100))
print("")
print("Interest Rate Range = {:.2f}% - {:.2f}%".format(int_rate_low * 100, int_rate_high * 100))
print("Interest Weighted Average = {:.2f}%".format(int_rate_average * 100))
print("")
print("Expected Term Range = {:.2f} years - {:.2f} years".format(exp_term_low, exp_term_high))
print("Expected Term Weighted Average = {:.2f}".format(exp_term_average))
print("")
print("Dividend Range = {}% - {}%".format(dividend_low, dividend_high))
print("Dividend Weighted Average = {}%".format(dividend_average))
print("")
print("Black-Scholes FMV Range = ${:.4f} - ${:.4f}".format(fair_market_low, fair_market_high))
print("Black-Scholes FMV Weighted Average = ${:.4f}".format(fair_market_average))


# Calculates the values required by the Option Activity section of disclosure #1
float_pre_total_outstanding = [float(i) for i in pre_total_outstanding]
float_pre_total_exercised = [float(i) for i in pre_total_exercised]
float_pre_total_forfeited = [float(i) for i in pre_total_forfeited]
pre_outstanding = sum(float_pre_total_outstanding)
pre_outstanding -= sum(float_pre_total_exercised)
pre_outstanding -= sum(float_pre_total_forfeited)
float_period_grants = [float(i) for i in period_total_grants]
float_outstanding_list = [float(i) for i in total_outstanding_list]
float_period_exercised = [float(i) for i in period_total_exercised]
float_period_forfeited = [float(i) for i in period_total_forfeited]
float_period_expired = [float(i) for i in period_expirations]
float_period_outstanding = [float(i) for i in period_outstanding]
float_end_period_exercisable = [int(i) for i in end_period_exercisable]
float_total_vested_list = [int(i) for i in total_vested_list]
total_outstanding_shares = pre_outstanding + sum(float_period_outstanding)
total_period_grants = float(sum(float_period_grants))
print("ATTENTION total_period_grants = {}".format(total_period_grants))

print(float_end_period_exercisable)

print("")
print("Total Outstanding (by period start) = {:.0f}".format(pre_outstanding))
print("")
print("Grants during the period = {:.0f}".format(sum(float_period_grants)))
print("")
print("Exercises during the period = {:.0f}".format(sum(float_period_exercised)))
print("")
print("Forfeitures during the period = {:.0f}".format(sum(float_period_forfeited)))
print("")
print("Expirations during the period = {:.0f}".format(sum(period_expirations)))
print("")
print("Total Outstanding at end of the period = {:.0f}".format(pre_outstanding + sum(float_period_outstanding)))
print("")
print("Total Exercisable at end of the period = {:.0f}".format(sum(float_end_period_exercisable)))
print("")
print("Total Unvested at Period End = {:.0f}".format(total_outstanding_shares - sum(float_total_vested_list)))
print("")
print("Total Vested = {:.0f}".format(sum(float_total_vested_list)))
print("")
print("Total Unrecognized Compensation = NOT OPERATIONAL")
print("")
print("Weighted Average Time to Recognize Unrecognized Compensation = NOT OPERATIONAL")

# This is the section that will export the Section #1 disclosure to an Excel worksheet
company_name = "SkyCraft Airplanes, Inc."
excel_file_name = "development.xlsx"

# company_name = input("Please enter your company name: ")
# excel_name = input("Please enter Excel file save name: ")
# excel_file_name = "{}.xlsx".format(excel_name)

# reporting_start_year = float(input("Please enter the year in which the reporting period starts> "))
# reporting_start_month = float(input("Please enter the month in which the reporting period starts> "))
# reporting_end_year = float(input("Please enter the length of your reporting period in months> "))


# Builds the static portion of the Excel Disclosure file
workbook = xlsxwriter.Workbook(excel_file_name)
worksheet1 = workbook.add_worksheet('Disclosure #1')
worksheet2 = workbook.add_worksheet('Disclosure #2')

# Adds types of formats included in the Excel file
background = workbook.add_format({'bg_color': '#858585'})
background_white = workbook.add_format({'bg_color': 'white'})
title_center = workbook.add_format({'bold': True, 'font_size': 18, 'align': 'center', 'bg_color': 'white'})
center = workbook.add_format({'align': 'center', 'bg_color': 'white'})
title_bottom_border = workbook.add_format({'bold': True, 'font_size': 18, 'bottom': 2, 'bg_color': 'white'})
headline1 = workbook.add_format({'bold': True, 'font_size': 12, 'font_color': '#007799',
                                 'bottom': 2, 'bottom_color': '#007799', 'bg_color': 'white'})
blue = workbook.add_format({'font_color': '#007799', 'bg_color': 'white'})
bold = workbook.add_format({"bold": True, 'bg_color': 'white'})
money = workbook.add_format({"num_format": "$#,##0", 'bg_color': 'white'})
date1 = workbook.add_format({'num_format': 'mm/dd/yyyy', 'align': 'center', 'bottom': 2, 'bg_color': 'white'})
bottom_border_right = workbook.add_format({'bottom': 2, 'bg_color': 'white', 'align': 'right'})
blue_bottom_thick = workbook.add_format({'bottom': 2, 'bottom_color': '#007799', 'bg_color': 'white'})
blue_bottom_thin = workbook.add_format({'bottom': 1, 'bottom_color': '#007799', 'bg_color': 'white', 'align': 'right'})
right_align = workbook.add_format({'align': 'right', 'bg_color': 'white'})
bottom_align = workbook.add_format({'align': 'bottom'})


# Set column width and background color
worksheet1.set_column('B:F', 10, background_white)
worksheet1.set_column('A:A', 15, background)
worksheet1.set_column('B:C', 30, background_white)
worksheet1.set_column('D:D', 40, background_white)
worksheet1.set_column('E:F', 30, background_white)
worksheet1.set_column('G:G', 15, background)
worksheet1.write('B1', "", background)
worksheet1.write('C1', "", background)
worksheet1.write('D1', "", background)
worksheet1.write('E1', "", background)
worksheet1.write('F1', "", background)
worksheet1.set_row(8, 18)

# Set borders for empty cells
worksheet1.write('C5', "", title_bottom_border)
worksheet1.write('D5', "", title_bottom_border)
worksheet1.write('C7', "", blue_bottom_thick)
worksheet1.write('D7', "", blue_bottom_thick)
worksheet1.write('E7', "", blue_bottom_thick)
worksheet1.write('F7', "", blue_bottom_thick)
worksheet1.write('B8', "", blue_bottom_thin)
worksheet1.write('E8', "", blue_bottom_thin)
worksheet1.write('F8', "", blue_bottom_thin)
worksheet1.write('C15', "", blue_bottom_thick)
worksheet1.write('D15', "", blue_bottom_thick)
worksheet1.write('E15', "", blue_bottom_thick)
worksheet1.write('F15', "", blue_bottom_thick)
worksheet1.write('B16', "", blue_bottom_thin)
worksheet1.write('B22', "", blue_bottom_thin)
worksheet1.write('C22', "", blue_bottom_thin)
worksheet1.write('D22', "", blue_bottom_thin)
worksheet1.write('E22', "", blue_bottom_thin)
worksheet1.write('F22', "", blue_bottom_thin)
worksheet1.write('C28', "", blue_bottom_thick)
worksheet1.write('D28', "", blue_bottom_thick)
worksheet1.write('E28', "", blue_bottom_thick)
worksheet1.write('F28', "", blue_bottom_thick)


# Write all static information
worksheet1.write('D2', 'Valuation Disclosure', title_center)
worksheet1.write('D3', 'Reporting Period: {} - {}'.format(excel_start_date, excel_end_date), center)
worksheet1.write('B5', company_name, title_bottom_border)
worksheet1.write('E5', 'Report Date:', bottom_border_right)
worksheet1.write('F5', '=TODAY()', date1)
worksheet1.write('B7', 'Valuation Summary', headline1)
worksheet1.write('C8', 'Range', blue_bottom_thin)
worksheet1.write('D8', 'Weighted Average', blue_bottom_thin)
worksheet1.write('B9', 'Volatility')
worksheet1.write('B10', 'Interest Rate')
worksheet1.write('B11', 'Expected Term')
worksheet1.write('B12', 'Dividend Rate')
worksheet1.write('B13', 'Fair Value Per Share on Grant Date')
worksheet1.write('B15', 'Option Activity', headline1)
worksheet1.write('C16', 'Shares', blue_bottom_thin)
worksheet1.write('D16', 'Exercise Weighted Price', blue_bottom_thin)
worksheet1.write('E16', 'Remaining Contract Average', blue_bottom_thin)
worksheet1.write('F16', 'Intrinsic Value', blue_bottom_thin)
worksheet1.write('B17', 'Total Outstanding (by period start)')
worksheet1.write('B18', 'Grants (during the period)')
worksheet1.write('B19', 'Exercises (during the period)')
worksheet1.write('B20', 'Forfeitures (during the period)')
worksheet1.write('B21', 'Expirations (during the period)')
worksheet1.write('B23', 'Total Outstanding (by period end)')
worksheet1.write('B24', 'Total Exercisable (by period end)')
worksheet1.write('B25', 'Total Unvested (by period end)')
worksheet1.write('B26', 'Total Vested (by period end)')
worksheet1.write('B28', 'Unrecognized Compensation', headline1)
worksheet1.write('B29', 'Total Unrecognized Compensation')
worksheet1.write('B30', 'Weighted Average Time to Recognize Unrecognized Compensation')


# Write all the dynamic information.
worksheet1.write('C9', '{}% - {}%'.format(vol_low * 100, vol_high * 100), right_align)
worksheet1.write('D9', '{:.2f}%'.format(vol_average * 100), right_align)
worksheet1.write('C10', '{:.2f}% - {:.2f}%'.format(int_rate_low * 100, int_rate_high * 100), right_align)
worksheet1.write('D10', '{:.2f}%'.format(int_rate_average * 100), right_align)
worksheet1.write('C11', '{:.2f} years - {:.2f} years'.format(exp_term_low, exp_term_high), right_align)
worksheet1.write('D11', '{:.2f} years'.format(exp_term_average), right_align)
worksheet1.write('C12', '{}% - {}%'.format(dividend_low, dividend_high), right_align)
worksheet1.write('D12', '{}%'.format(dividend_average), right_align)
worksheet1.write('C13', '${:.4f} - ${:.4f}'.format(fair_market_low, fair_market_high), right_align)
worksheet1.write('D13', '${:.4f}'.format(fair_market_average), right_align)
worksheet1.write('C17', '{:.0f}'.format(pre_outstanding), right_align)
worksheet1.write('C23', '{:.0f}'.format(pre_outstanding + sum(float_period_outstanding)), right_align)
worksheet1.write('D23', '${:.4f}'.format((weighted_grants + weighted_outstanding) /
                                         (pre_outstanding + sum(float_period_grants))), right_align)

# In the event that "Total Outstanding Shares at Period Start" is 0, this changes the value to 1 so we don't get
# an error when trying to divide by 0. If this value is 0 the value we are dividing will also be zero. So effectively
# we get the same result without the error.
if pre_outstanding == 0:
    pre_outstanding = 1

worksheet1.write('D17', '${:.4f}'.format(weighted_outstanding / pre_outstanding), right_align)
worksheet1.write('C18', '{:.0f}'.format(sum(float_period_grants)), right_align)
worksheet1.write('D18', '${:.4f}'.format(weighted_grants / sum(float_period_grants)), right_align)
worksheet1.write('C19', '{:.0f}'.format(sum(float_period_exercised)), right_align)
worksheet1.write('D19', '${:.4f}'.format(weighted_exercised / sum(float_period_exercised)), right_align)
worksheet1.write('C20', '{:.0f}'.format(sum(float_period_forfeited)), right_align)
worksheet1.write('D20', '${:.4f}'.format(weighted_forfeitures / sum(float_period_forfeited)), right_align)
worksheet1.write('C21', '{:.0f}'.format(sum(float_period_expired)), right_align)

# In the event that "Expirations during the period" is 0, this changes the value to 1 so we don't get
# an error when trying to divide by 0. If this value is 0 the value we are dividing will also be zero. So effectively
# we get the same result without the error.
if sum(float_period_expired) == 0:
    worksheet1.write('D21', '$0.0000', right_align)
else:
    worksheet1.write('D21', '${:.4f}'.format(weighted_expirations / sum(float_period_expired)), right_align)

worksheet1.write('C24', '{:.0f}'.format(sum(float_end_period_exercisable)), right_align)
worksheet1.write('D24', '${:.4f}'.format(weighted_exercisable / sum(float_end_period_exercisable)), right_align)
worksheet1.write('C25', '{:.0f}'.format(total_unvested), right_align)
worksheet1.write('D25', '${:.4f}'.format(weighted_unvested / total_unvested), right_align)
worksheet1.write('C26', '{:.0f}'.format(sum(float_total_vested_list)), right_align)
worksheet1.write('D26', '${:.4f}'.format(weighted_vested / sum(float_total_vested_list)), right_align)
worksheet1.write('E26', '{} years'.format(contract_term / sum(float_total_vested_list)), right_align)

# Adds the Shoobx logo to the Excel file
worksheet1.insert_image('B2', 'shoobx_logo.png', {'x_scale': 1.0, 'y_scale': 0.8, 'x_offset': 10, 'y_offset': 10})





workbook.close()

f.close()
