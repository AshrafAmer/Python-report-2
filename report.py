import xlrd
import datetime
import math


def getting_precent( value ):
    """
    """
    if ( math.ceil(value) > (value + 0.5) ):
    	return( int(value *100) )
    else:
    	return( math.ceil(value)*100 ) 


def positive_trades( value ):
    """
    """
    count = 0
    for i in value:
    	if i > 0:
    		count += 1
    return count


def open_file(path):
    """
    #######################################
    #######################################
    """

    # open file
    book = xlrd.open_workbook(path)
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)

    # output reading data array
    # read all trades of spread which equal to orange col. 
    trades_of_spread = []
    trades_of_bean = []
    trades_of_meal = []
    trades_of_oil = []
    # all negative cols in [U].
    all_neg = []
    # read excel file
    for x in range(3, 15708):
        # read spread [AD]
        read_spread = first_sheet.col_values(29)[x]
        if (read_spread != ''):
            trades_of_spread.append(read_spread)
        # read bean [AA]
        read_bean = first_sheet.col_values(26)[x]
        if (read_bean != ''):
            trades_of_bean.append(read_bean)
        # read meal [AB]
        read_meal = first_sheet.col_values(27)[x]
        if (read_meal != ''):
            trades_of_meal.append(read_meal)
        # read oil [AC]
        read_oil = first_sheet.col_values(28)[x]
        if (read_oil != ''):
           trades_of_oil.append(read_oil)
        # read all neagative on col.[U]=> orange and green
        sl_bean = first_sheet.col_values(20)[x]
        if (sl_bean != '' and int(sl_bean) < 0):
            all_neg.append(sl_bean)
        


    # Prepare Data for Percent Profitable
    positive_spread = positive_trades(trades_of_spread)
    positive_bean = positive_trades(trades_of_bean)
    positive_meal = positive_trades(trades_of_meal)
    positive_oil = positive_trades(trades_of_oil)
    len_spread = len(trades_of_spread)
    len_all_neg = len(all_neg)


    # Output table.
    print( "\t\t  ", "[ Spread, Bean, Meal, Oil ]:" )
    # Number of trades col.
    print( "Number of Trades  :[ ", len_spread, ', ',
    	len(trades_of_bean), ', ', len(trades_of_meal), ', ',
    	len(trades_of_oil),' ]' )
    # Percent Profitable
    print("Percent Profitable:[ ",
    	getting_precent((positive_spread/len_spread)),'%, ',
    	getting_precent((positive_bean/len(trades_of_bean))),'%, ',
    	getting_precent((positive_meal/len(trades_of_meal))),'%, ',
    	getting_precent((positive_oil/len(trades_of_oil))),'% ]'
    	)
    # Maximum porfit
    print( "Maximum Porfit    :[ $",
    	("%.2f" %max(trades_of_spread)), ' ,$',
    	("%.2f" %max(trades_of_bean)), ' ,$',
    	("%.2f" %max(trades_of_meal)), ' ,$',
    	("%.2f" %max(trades_of_oil)), ' ]' )
    # Minimum porfit
    print( "Minimum Porfit    :[ $",
    	("%.2f" %min(trades_of_spread)), ' ,$',
    	("%.2f" %min(trades_of_bean)), ' ,$',
    	("%.2f" %min(trades_of_meal)), ' ,$',
    	("%.2f" %min(trades_of_oil)), ' ]' )
    # Average porfit
    print( "Average Porfit    :[ ",
    	("%.2f" %(sum(trades_of_spread)/len_spread)), ' ,',
    	("%.2f" %(sum(trades_of_bean)/len(trades_of_bean))), ' ,',
    	("%.2f" %(sum(trades_of_meal)/len(trades_of_meal))), ' ,',
    	("%.2f" %(sum(trades_of_oil)/len(trades_of_oil))), ' ]' )

    # print contango and backwardation
    print("All Number of Episodes - Contango:[ ",
    	len_spread, ' ,', len_spread, ' ,',
    	len_spread, ' ,', len_spread, ' ]')

    val = len_all_neg - len_spread
    print("All Number of Episodes - Backwardation:[ ",
    	val, ' ,', val, ' ,', val, ' ,', val, ' ]')
 
  
  
if __name__ == "__main__":
    path = "test.xls"
    open_file('[File-Name].xlsx')