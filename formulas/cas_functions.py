# -*- coding: utf-8 -*- 
import datetime
from datetime import date
from formulas.helpers.numbers import *
from config import config as cfg

class CasFunctions():
    config = cfg.Config()
    grp_baza = config.grp_baza
    e2_baza = config.e2_baza
    cas_baza = config.cas_baza
    shr_baza = config.shr_baza
    customers = config.customers()
    sk_nova = config.sk_nova()
    cas_collection = config.cas_collection()


    cas_others_dm_pool = config.cas_others_dm_pool
    cas_non_agb = config.cas_non_agb
    cas_others = config.cas_others
    cas_dm_pool_channels = config.cas_dm_pool_channels
    dm_pool_shr_channels = config.dm_pool_shr_channels
    cas_other_channels = config.cas_other_channels

    range_by_quarter = {
        1: 0,
        2: 3,
        3: 6,
        4: 9
    }

    duration_by_quarter = {
        1: 1,
        2: 5,
        3: 9,
        4: 13
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

    marko_shr_total = {
        1: [0,1,2,3,4,17],
        2: [0,4,5,6,7,8,17],
        3: [0,4,8,9,10,11,12,17],
        4: [0,4,8,12,13,14,15,16,17]
    }


    def __init__(self,year,quarter,collection=None):
        self.collection = collection
        self.CURRENT_YEAR = year
        self.QUARTER = quarter

    def channel_total_ly(self,customer,channel):
        t = customer
        c = channel
        return sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == t and channel == c])

    def channel_total_ly_no_customer(self,channel):
        c = channel
        return sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and channel == c])

    # zbir kanala za proslu godinu
    def total_customer_ly(self):
        res = {}
        for cas in self.customers:
            total = sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == cas and channel in self.cas_other_channels])
            res[cas] = total
        return res
      
    def cas_others_total_customer_ly(self):
        res = {}
        for cas in self.customers:
            total = sum([total for (total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == cas and channel in self.cas_others_dm_pool])
            res[cas] = total
        return res


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
        return sum([total for (total,year,month,customer,channel) in self.collection if year == y and month == m and customer == t and channel in self.cas_dm_pool_channels])

    # matrica za trenutnu god za sve kanale
    def channels_for_customer2(self):
        print("RACUNAM SVE KANALE!")
        res = {}
        for cas in self.customers:
            for c in self.cas_other_channels:
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

    def total_amount_for_months(self,total_for_one_month):
        res = {}
        for m in range(1,13):
            total = 0
            for g in total_for_one_month.values():
                total += g[m-1]
            res[m] = total
        return res

    def n1_nova_s_total_for_months(self,all_channels):
        res = {}
        for ch in ['N1','Nova S']:
            f = {}
            for m in range(0,12):
                total = 0
                for k,v in all_channels.items():
                    channel = k.split('_')[1]
                    if ch == channel:
                        total += v[m]
                f[m] = total
            res[ch] = f
        return res

    def cas_others_total_for_months(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            for m in range(0,12):
                total = 0
                for k,v in all_channels.items():
                    if cas in k and 'N1' not in k and 'Nova S' not in k:
                        total += v[m]
                f[m] = total
            res[cas] = f
        return res

    def n1_nova_s_channel_total_ly(self):
        res = {}
        for ch in ['N1','Nova S']:
            res[ch] = self.channel_total_ly_no_customer(ch)
        return res


    ##################
    ##### SHR ########
    ##################

    def shr_channels(self):
        ch = self.dm_pool_shr_channels.copy()
        ch.remove('Sport klub SRB')
        ch.remove('CAS "non AGB" (SK, Nova Sport, Brainz)')
        ch.remove('CAS Others (w/o SK, Nova Sport, Brainz)')
        ch += ['SK 1', 'SK 2', 'SK 3']
        return ch 

    def shr_collection(self):
        values = {}
        for row in self.shr_baza.rows:
            res = None
            if row[0].value in self.shr_channels():
                res = []
                tap = ()
                for cell in row:
                    if cell.value == None or cell.value == ' ':
                        val = 0
                    else:
                        val = cell.value
                    tap = tap + (val,)
                res.append(tap)
            if res != None:
              values[row[0].value] = res
        return values

    def shr_value(self,channel,column):
        coll = self.shr_collection().get(channel)
        sum = 0
        if coll != None:
          for c in coll:
              sum = sum + c[column]
        return (sum * 100)

    def formated_shr_value(self,channel,column):
        val = self.shr_value(channel,column)
        return "{0}%".format(round(val, 2))

    def shr_channel_value(self,m1,m2):
        arr = []
        for c in self.shr_channels():
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.formated_shr_value(c,i),)
            arr.append(list(tup))
        self.merge_sport_klub(arr)
        self.cas_others_total(arr)
        self.cas_non_agb_total(arr)
        return arr

    def non_formated_shr_value(self,channel,column):
        val = self.shr_value(channel,column)
        return "{0}%".format(val)

    def non_shr_channel_value(self,m1,m2):
        arr = []
        for c in self.shr_channels():
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.non_formated_shr_value(c,i),)
            arr.append(list(tup))
        self.non_merge_sport_klub(arr)
        self.non_cas_others_total(arr)
        self.non_cas_non_agb_total(arr)
        return arr

    def non_shr_channel_output(self,m1,m2):
        arr = []
        for c in self.shr_channels():
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.non_formated_shr_value(c,i),)
            arr.append(list(tup))
        return arr

    def non_shr_channel_value_2(self,m1,m2):
        arr = []
        for c in self.shr_channels():
            tup = (c,)
            for i in list(range(m1,m2)):
                tup = tup + (self.non_formated_shr_value_2(c,i),)
            arr.append(list(tup))
        arr2 = self.non_shr_channel_output(m1,m2)
        a1 = self.non_merge_sport_klub_2(arr2)
        a2 = self.non_cas_others_total_2(arr2)
        a3 = self.non_cas_non_agb_total_2(arr2)

        arr.append(a1)
        arr.insert(2,a2)
        arr.insert(-6,a3)
        return arr

    def non_formated_shr_value_2(self,channel,column):
        val = self.shr_value(channel,column)
        return "{0}%".format(round(val, 2))

    def non_merge_sport_klub(self,res):
        new_arr = ['Sport klub SRB']
        s = [ (f) for (ch,*f) in res if ch in ['SK 1','SK 2', 'SK 3']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(total))
        res.append(new_arr)
        return res

    def non_cas_others_total(self,res):
        new_arr = ['CAS Others (w/o SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch not in ['Brainz','Nova Sport','Sport klub SRB','N1','Nova S','SK 1','SK 2', 'SK 3','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(total))
        res.insert(2,new_arr)
        return res

    def non_cas_non_agb_total(self,res):
        new_arr = ['CAS "non AGB" (SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch in ['Brainz','Nova Sport','Sport klub SRB']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(total))
        res.insert(-6,new_arr)
        return res

    # V2

    def non_merge_sport_klub_2(self,res):
        new_arr = ['Sport klub SRB']
        s = [ (f) for (ch,*f) in res if ch in ['SK 1','SK 2', 'SK 3']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        return new_arr

    def non_cas_others_total_2(self,res):
        new_arr = ['CAS Others (w/o SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch not in ['Brainz','Nova Sport','Sport klub SRB','N1','Nova S','SK 1','SK 2', 'SK 3','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        # res.insert(2,new_arr)
        return new_arr

    def non_cas_non_agb_total_2(self,res):
        new_arr = ['CAS "non AGB" (SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch in ['Brainz','Nova Sport','SK 1','SK 2', 'SK 3']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        # res.insert(-6,new_arr)
        return new_arr

    def merge_sport_klub(self,res):
        new_arr = ['Sport klub SRB']
        s = [ (f) for (ch,*f) in res if ch in ['SK 1','SK 2', 'SK 3']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        res.append(new_arr)
        return res

    def cas_non_agb_total(self,res):
        new_arr = ['CAS "non AGB" (SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch in ['Brainz','Nova Sport','Sport klub SRB']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        res.insert(-6,new_arr)
        return res

    def cas_others_total(self,res):
        new_arr = ['CAS Others (w/o SK, Nova Sport, Brainz)']
        s = [ (f) for (ch,*f) in res if ch not in ['Brainz','Nova Sport','Sport klub SRB','N1','Nova S','SK 1','SK 2', 'SK 3','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            new_arr.insert(i+1,'{0}%'.format(round(total,2)))
        res.insert(2,new_arr)
        return res


    ########################################
    ############# DURATION #################
    ########################################

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


    def duration_total(self,year,month):
        y = year
        m = month
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == y and month == m]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

    # matrica za trenutnu god za sve kanale
    def duration_channel_total(self):
        res = {}
        for cas in self.customers:
            for c in self.cas_other_channels:
                channel_row = ()
                for m in range(0,12):
                    val = self.duration_channel(self.CURRENT_YEAR,m+1,cas,c)
                    channel_row = channel_row + (val,)
                res["{0}_{1}".format(cas,c)] = channel_row
        return res

    def duration_total_by_month(self):
        channel_row = ()
        for m in range(0,12):
            val = self.duration_total(self.CURRENT_YEAR,m+1)
            channel_row = channel_row + (val,)
        return channel_row

    def cas_duration(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            for m in range(0,12):
                total = datetime.timedelta()
                for k,v in all_channels.items():
                    if cas in k:
                        total += v[m]
                f[m] = total
            res[cas] = f
        return res


    def total_duration_for_months(self,total_for_one_month):
        res = {}
        for m in range(1,13):
            total = datetime.timedelta()
            for g in total_for_one_month.values():
                total += g[m-1]
            res[m] = total
        return res

    def n1_nova_s_duration(self,all_channels):
        res = {}
        for ch in ['N1','Nova S']:
            f = {}
            for m in range(0,12):
                total = datetime.timedelta()
                for k,v in all_channels.items():
                    channel = k.split('_')[1]
                    if ch == channel:
                        total += v[m]
                f[m] = total
            res[ch] = f
        return res

    def cas_others_duration(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            for m in range(0,12):
                total = datetime.timedelta()
                for k,v in all_channels.items():
                    if cas in k and 'N1' not in k and 'Nova S' not in k:
                        total += v[m]
                f[m] = total
            res[cas] = f
        return res

    def sum_delta_time(self,arr):
        mysum = datetime.timedelta()
        for i in arr:
            mysum += i
        return mysum

    def last_year_duration_channel_total(self):
        res = {}
        for cas in self.customers:
            for c in self.cas_other_channels:
                val = self.duration_customer_channel_ly(cas,c)
                res["{0}_{1}".format(cas,c)] = val
        return res

    def last_year_cas_duration(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            total = datetime.timedelta()
            for k,v in all_channels.items():
                if cas in k:
                    total += v
            res[cas] = total
        return res

    def cas_others_duration_last_year(self,all_channels):
        res = {}
        for cas in self.customers:
            f = {}
            total = datetime.timedelta()
            for k,v in all_channels.items():
                if cas in k and 'N1' not in k and 'Nova S' not in k:
                    total += v
            res[cas] = total
        return res

    def duration_all_customers_total_ly(self):
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

    def total_duration_for_months_last_year(self,total_for_one_month):
        res = {}
        total = datetime.timedelta()
        for g in total_for_one_month.values():
            total += g
        return total

    def n1_nova_s_duration_last_year(self,all_channels):
        res = {}
        for ch in ['N1','Nova S']:
            f = {}
            total = datetime.timedelta()
            for k,v in all_channels.items():
                channel = k.split('_')[1]
                if ch == channel:
                    total += v
            res[ch] = total
        return res

    def duration_customer_channel_ly(self,customer,channel):
        t = customer
        c = channel
        total_time_arr = [total for (*total,year,month,customer,channel) in self.collection if year == self.CURRENT_YEAR-1 and customer == t and channel == c]
        flat_list = [item for tta in total_time_arr for item in tta]
        return self.calculate_timedelta(flat_list)

  #########################
  ##### SOI ##############
  ########################


    def soi_e2_channels(self):
        ch = self.dm_pool_shr_channels.copy()
        ch.remove('CAS "non AGB" (SK, Nova Sport, Brainz)')
        ch.remove('CAS Others (w/o SK, Nova Sport, Brainz)')
        return ch 

    def soi_cas_channel_total(self,month,channel):
        m = month
        ch = channel
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR and month == [m] and channel == ch])

  
    def soi_total_ly(self,channel):
        ch = channel
        return sum([total for (total,year,*month,channel) in self.cas_collection if year == self.CURRENT_YEAR-1 and channel == ch])


    def soi_channel_total_ly(self):
        res = {}
        for c in self.soi_e2_channels():
            val = self.soi_total_ly(c)
            res[c] = val
        return res

    def soi_cas_others_channels_ly(self,all_channels):
        total = 0
        for k,v in all_channels.items():
            if k not in ['N1','Nova S'] + list(self.sk_nova):
                total += v
        return total

    def cas_media_agb_ly(self,all_channels):
        total = 0
        for k,v in all_channels.items():
            if k not in self.sk_nova:
                total += v
        return total

    def cas_media_channels_ly(self,all_channels):
        total = 0
        for k,v in all_channels.items():
            if k in self.sk_nova:
                total += v
        return total

    # matrica za trenutnu god za sve kanale
    def e2_channel_total(self):
        res = {}
        for c in self.soi_e2_channels():
            channel_row = ()
            for m in range(0,12):
                val = self.soi_cas_channel_total(m+1,c)
                channel_row = channel_row + (val,)
            res[c] = channel_row
        return res

    # prvi red
    def cas_media_total(self,all_channels):
        res = {}
        for m in range(0,12):
            total = 0
            for k,v in all_channels.items():
                total += v[m]
            res[m] = total
        return res

    # drugi red
    def drugi_red_total(self,all_channels):
        res = {}
        for m in range(0,12):
            total = 0
            for k,v in all_channels.items():
                if k not in self.sk_nova:
                    total += v[m]
            res[m] = total
        return res

    # poslednje 3
    def cas_media_channels_agb(self,all_channels):
        res = {}
        for ch in self.sk_nova:
            f = {}
            for m in range(0,12):
                total = 0
                for k,v in all_channels.items():
                    if ch == k:
                        total += v[m]
                f[m] = total
            res[ch] = f
        return res


    # CAS U Kanalima
    def soi_cas_others(self,all_channels):
        res = {}
        for m in range(0,12):
            total = 0
            for k,v in all_channels.items():
                if k not in ['N1','Nova S'] + list(self.sk_nova):
                    total += v[m]
            res[m] = total
        return res

    # AGB u kanalima
    def soi_agb_channal(self,all_channels):
        res = {}
        for m in range(0,12):
            total = 0
            for k,v in all_channels.items():
                if k in self.sk_nova:
                    total += v[m]
            res[m] = total
        return res


      
