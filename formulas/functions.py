# -*- coding: utf-8 -*- 
import datetime
from datetime import date
from formulas.helpers.numbers import *
from config import config as cfg


class Functions():
    config = cfg.Config()
    grp_baza = config.grp_baza
    e2_baza = config.e2_baza
    cas_baza = config.cas_baza
    shr_baza = config.shr_baza
    customers = config.customers()
    channels = config.channels()
    e2_collection = config.e2_collection()
    cas_collection = config.cas_collection()
    e2_channels = config.e2_channels()
    cas_channels = config.cas_channels()
    sk_nova = config.sk_nova()
    channel_map = config.channel_map

    range_by_quarter = {
        1: 0,
        2: 3,
        3: 6,
        4: 9
    }

    shr_total_population_range = {
        1: [13,14,15,16,17,30],
        2: [13,17,18,19,20,21,30],
        3: [13,17,21,22,23,24,25,30],
        4: [13,17,21,25,26,27,28,29,30]
    }

    shr_total_18_50_range = {
        1: [45,46,47,48,49,62],
        2: [45,49,50,51,52,53,62],
        3: [45,49,53,54,55,56,57,62],
        4: [45,49,53,57,58,59,60,61,62]
    }


    def __init__(self,year,quarter,collection=None):
        self.collection = collection
        self.CURRENT_YEAR = year
        self.QUARTER = quarter

    ###################### NOVE FUNKCIJE #####################
    def channel_total_ly(self,customer,channel):
        t = customer
        c = channel
        return sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == t and channel == c])

    # zbir kanala za proslu godinu
    def total_customer_ly(self,customer):
        t = customer
        return sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == t])

    ############# KRAJ PROSLA GOD ############
    def channel_total(self,year,month,customer,channel):    
        y = year
        m = month
        t = customer
        c = channel
        return sum([total for (total,year,month,customer,channel) in self.collection if year == y and month == m and customer == t and channel == c])

    def customer_total(self,year,month,customer,channel):
        y = year
        m = month
        t = customer
        return sum([total for (total,year,month,customer,channel) in self.collection if year == y and month == m and customer == t])

    def channels_for_customer(self,customer):
        res = {}
        cas = customer
        for c in self.channels:
            channel_row = ()
            for m in range(1,13):
                val = self.channel_total(self.CURRENT_YEAR,m,cas,c)
                channel_row = channel_row + (val,)
            res["{0}_{1}".format(cas,c)] = channel_row
        return res

    def total_customer_for_month(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            for m in range(0,12):
                total = 0
                for k,v in all_channels.items():
                    if cas in k:
                        total += v[m]
                f[m] = total
            res[cas] = f
        return res

    # matrica za trenutnu god za sve kanale
    def channels_for_customer2(self):
        print("RACUNAM SVE KANALE!")
        res = {}
        for cas in self.customers:
            for c in self.channels:
                channel_row = ()
                for m in range(1,13):
                    val = self.channel_total(self.CURRENT_YEAR,m,cas,c)
                    channel_row = channel_row + (val,)
                res["{0}_{1}".format(cas,c)] = channel_row
        return res

    
    def find_by_channel(self,customer,channel):
        print("Racuna {0} ---> {1}".format(customer,channel))
        s = self.channels_for_customer(customer)
        return s["{0}_{1}".format(customer,channel)]


    # quartalni za customer (quarter > 4 - za celu god)
    def total_customer_for_quarter(self,customer):
        print("Quarter za customer: {0}".format(customer))
        a = self.channels_for_customer(customer)
        hash = {}
        for i in range(1,6):
            total = 0
            for k,v in a.items():
                if i == 1:
                    total = total + v[0] + v[1] + v[2]
                elif i == 2:
                    total = total + v[3] + v[4] + v[5]
                elif i == 3:
                    total = total + v[6] + v[7] + v[8]
                elif i == 4:
                    total = total + v[9] + v[10] + v[11]
                else:
                    total = total + sum(v)
            hash[i] = total
        return hash

    def all_customers_total(self,year,month=None):
        y = year
        m = month
        if m != None:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == y and month == m]))
        else:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == y]))

    def all_customers_total_for_quarter(self,quarter):
        if quarter == 1:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR and month in (1,2,3)]))
        elif quarter == 2:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR and month in (4,5,6)]))
        elif quarter == 3:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR and month in (7,8,9)]))
        else:
            return round(sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR and month in (10,11,12)]))


    def total_amount_for_months(self,total_for_one_month):
        res = {}
        for m in range(1,13):
            total = 0
            for g in total_for_one_month.values():
                total += g[m-1]
            res[m] = total
        return res

    def grp_channels_by_customer(self):
        return self.channels_for_customer2()


    def grp_customers_quarter(self,total_for_one_month):
        res = {}
        for cas in self.customers:
            res[cas] = self.grp_customers_quarter2(total_for_one_month[cas])
        return res

    def grp_customers_quarter2(self,v):
        hash = {}
        for i in range(1,6):
            total = 0
            if i == 1:
                total = total + v[0] + v[1] + v[2]
            elif i == 2:
                total = total + v[3] + v[4] + v[5]
            elif i == 3:
                total = total + v[6] + v[7] + v[8]
            elif i == 4:
                total = total + v[9] + v[10] + v[11]
            else:
                total = sum(v.values())
            hash[i] = total
        return hash

    def shr_collection(self):
        values = {}
        for c in self.channels:
            res = []
            for row in self.shr_baza.rows:
                if row[0].value in self.channel_map[c]:
                    tap = ()
                    for cell in row:
                        if cell.value == None or cell.value == ' ':
                            val = 0
                        else:
                            val = cell.value
                        tap = tap + (val,)
                    res.append(tap)
            values[c] = res
        return values

    def shr_value(self,channel,column):
        coll = self.shr_collection()[channel]
        sum = 0
        for c in coll:
            sum = sum + c[column]
        return (sum * 100)

    def formated_shr_value(self,channel,column):
        val = self.shr_value(channel,column)
        return "{0}%".format(round(val, 2))
    
    def non_formated_shr_value(self,channel,column):
        val = self.shr_value(channel,column)
        return "{0}%".format(val)

    def shr_total_value(self,column):
        res = 0
        for c in self.channels:
            res = res + self.shr_value(c,column)
        return "{0}%".format(round(res,1))

    def shr_channel_value(self,m1,m2):
        arr = []
        for c in self.channels:
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.formated_shr_value(c,i),)
            arr.append(list(tup))
        return arr

    def non_shr_channel_value(self,m1,m2):
        arr = []
        for c in self.channels:
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.non_formated_shr_value(c,i),)
            arr.append(list(tup))
        return arr

    def dm_pool_shr_total(self,m1,m2):
        res = {}
        c = self.non_shr_channel_value(m1,m2)
        for i in range(1,19):
            total = 0
            for p in c:
                total = total + float(p[i][:-1])
            res[i] = '{0}%'.format(round(total,1))
        return res


    def calculate_timedelta(self,time):
        mysum = datetime.timedelta()
        for i in time:
            if type(i) == str:
                (h, m, s) = i.split(':')
                d = datetime.timedelta(hours=int(h), minutes=int(m), seconds=int(s))
                mysum += d
            else:
                mysum += i
        return mysum


    def duration_channel(self,year,month,customer,channel):
        y = year
        m = month
        t = customer
        c = channel
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == y and month == m and customer == t and channel == c]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)


    def duration_customer_total(self,year,month,customer):
        y = year
        m = month
        t = customer
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == y and month == m and customer == t]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

    def duration_customers_total(self):
        res = {}
        for cas in self.customers:
            channel_row = ()
            for m in range(0,12):
                val = self.duration_customer_total(self.CURRENT_YEAR,m+1,cas)
                channel_row += (val,)
            res[cas] = channel_row
        return res


    # matrica za trenutnu god za sve kanale
    def duration_channel_total(self):
        res = {}
        for cas in self.customers:
            for c in self.channels:
                channel_row = ()
                for m in range(0,12):
                    val = self.duration_channel(self.CURRENT_YEAR,m+1,cas,c)
                    channel_row = channel_row + (val,)
                res["{0}_{1}".format(cas,c)] = channel_row
        return res

    def total_duration_for_months(self):
        res = {}
        dctd = self.duration_customers_total()
        for m in range(0,12):
            total = datetime.timedelta()
            for k,v in dctd.items():
                total += v[m]
            res[m] = total
        return res


    def sum_delta_time(self,arr):
        mysum = datetime.timedelta()
        for i in arr:
            mysum += i
        return mysum

    def duration_customer_total_ly(self,customer):
        t = customer
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == t]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

    def duration_all_customers_total_ly(self):
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

    def duration_channel_total_ly(self,customer,channel):
        t = customer
        c = channel
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1  and customer == t and channel == c]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)


    #### PROSLA GOD
    def e2_channel_total_ly(self,channel):
        c = channel
        return sum([total for (total,year,*month,channel) in self.e2_collection if year == self.CURRENT_YEAR-1 and channel == c])

    def SOI_total_ly(self,coll):
        return sum([total for (total,year,*month,channel) in coll if year == self.CURRENT_YEAR-1 and channel not in self.cas_channels])

    def cas_media_ly(self):
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR-1 and channel not in self.sk_nova])

    def sk_nova_ly(self):
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR-1 and channel in self.sk_nova])

    def e2_channel_total(self,month,channel,coll):
        m = month
        c = channel
        return sum([total for (total,year,month,channel) in coll if year == self.CURRENT_YEAR and month == m and channel == c])

    # matrica za trenutnu god za sve kanale
    def e2_channels_value(self,channels,coll):
        res = {}
        for c in channels:
            channel_row = ()
            for m in range(1,13):
                val = self.e2_channel_total(m,c,coll)
                channel_row = channel_row + (val,)
            res[c] = channel_row
        return res


    def cas_media_total(self,month):
        m = month
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR and month == [m] and channel not in self.sk_nova])

    def sk_nova_total(self,month):
        m = month
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR and month == [m] and channel in self.sk_nova])

    def cas_sk_channel_value(self):
        res = {}
        cas_row = ()
        sk_row = ()
        for m in range(1,13):
            cas = self.cas_media_total(m)
            sk = self.sk_nova_total(m)
            cas_row = cas_row + (cas,)
            sk_row = sk_row + (sk,)
        res['Cas Media'] = cas_row
        res['CAS (SK, Nova Sport, Brainz)'] = sk_row
        return res

    def SOI_total_for_month(self):
        f = {}
        a = self.e2_channels_value(self.e2_channels,self.e2_collection)
        b = self.cas_sk_channel_value()
        for m in range(0,12):
            total = 0
            for k,v in a.items():
                total = total + v[m]
            for k,v in b.items():
                total = total + v[m]
            f[m] = total
        return f

    def total_sk_nova(self):
        t = 0
        for i in range(1,13):
            t = t + self.sk_nova_total(i)
        return t


    def cas_total_for_month(self):
        res = {}
        for m in range(0,12):
            total = 0
            for k,v in self.cas_sk_channel_value().items():
                total = total + v[m]
            res[m] = total
        return res

    def e2_cas_sk(self):
        return self.e2_channels_value(self.e2_channels,self.e2_collection) | self.cas_sk_channel_value()

