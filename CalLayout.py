
import math

def excel_floor(number, significance=1.0):
    return math.floor(number / significance) * significance


def CalLayout(len,lonSide_width,shoSide_width,bias=0.15):
    ITO_LEN = 398.40
    ITO_WID = 347.60

    Y_pro = (lonSide_width-bias) + (shoSide_width-bias)
    X_pro = len-bias


    cal_X01 = excel_floor(ITO_LEN/X_pro, 1)
    cal_Y01 = excel_floor(ITO_WID/Y_pro, 0.5)

    cal_Y02 = excel_floor(ITO_LEN/Y_pro, 0.5)
    cal_X02 = excel_floor(ITO_WID/X_pro, 1)


    CAL_01 = cal_X01 * cal_Y01 * 2
    CAL_02 = cal_Y02 * cal_X02 * 2

    return CAL_01,CAL_02




if __name__ == '__main__':
    #print(CalLayout(37.45,30.80,26.80))
    #a = CalLayout(52,21.2,19.2)
    #print()
    pass


