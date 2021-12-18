# -*- coding: utf-8 -*- 

from formulas.helpers.numbers import *
import formulas.functions as f
import style.cover as style


class Cover():

    def __init__(self,year,quarter,sources,collection=None):
        self.sources = sources
        self.collection = collection
        self.CURRENT_YEAR = year
        self.QUARTER = quarter
        
        if collection != None:
            self.fn = f.Functions(year,quarter,collection)
            if 'grp' in sources:
                self.grp_channels = self.fn.channels_for_customer2()
                self.grp_customers_v = self.fn.total_customer_for_month(self.grp_channels)
                self.total_for_customers = self.fn.total_amount_for_months(self.grp_customers_v)
                self.grp_customers_quarter_v = self.fn.grp_customers_quarter(self.grp_customers_v)
            if 'duration' in sources:
                self.c_duration = self.fn.duration_channel_total()
                self.cas_duration = self.fn.duration_customers_total()
                self.all_duration = self.fn.total_duration_for_months()
                self.duration_ly = self.fn.duration_ly
            if 'soi' in sources:
                self.soi_total = self.fn.SOI_total_for_month()
                self.e2_cas_sk = self.fn.e2_channels_value(self.fn.e2_channels,self.fn.e2_collection) | self.fn.cas_sk_channel_value()
                p = sorted(self.e2_cas_sk.keys())
                p.remove('Cas Media')
                p.insert(1,'Cas Media')
                self.all_c = p
                self.last_year_total = self.fn.SOI_total_ly(self.fn.e2_collection) + self.fn.SOI_total_ly(self.fn.cas_collection)
                self.soi_grp_channels = self.fn.channels_for_customer2()
                self.dm_pool = self.fn.total_customer_for_month(self.soi_grp_channels)['DM Pool']

        self.headers = ["{0} FY".format(self.CURRENT_YEAR-1), "{0} JAN".format(self.CURRENT_YEAR), "{0} FEB".format(self.CURRENT_YEAR), "{0} MAR".format(self.CURRENT_YEAR), "{0} Q1".format(self.CURRENT_YEAR), "{0} APR".format(self.CURRENT_YEAR), "{0} MAJ".format(self.CURRENT_YEAR), "{0} JUN".format(self.CURRENT_YEAR), "{0} Q2".format(self.CURRENT_YEAR), "{0} JUL".format(self.CURRENT_YEAR), "{0} AVG".format(self.CURRENT_YEAR), "{0} SEP".format(self.CURRENT_YEAR), "{0} Q3".format(self.CURRENT_YEAR), "{0} OKT".format(self.CURRENT_YEAR), "{0} NOV".format(self.CURRENT_YEAR), "{0} DEC".format(self.CURRENT_YEAR), "{0} Q4".format(self.CURRENT_YEAR), "{0} ytd".format(self.CURRENT_YEAR)]

    def total_table(self):
        data = [['Total']]
        data.append(['Row Labels']+self.headers)
        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            data.append([cas,float_with_comma(round(self.fn.total_customer_ly(cas))),float_with_comma(round(cas_total[0])),float_with_comma(round(cas_total[1])),float_with_comma(round(cas_total[2])),float_with_comma(round(cas_total_quarter[1])),float_with_comma(round(cas_total[3])),float_with_comma(round(cas_total[4])),float_with_comma(round(cas_total[5])),float_with_comma(round(cas_total_quarter[2])),float_with_comma(round(cas_total[6])),float_with_comma(round(cas_total[7])),float_with_comma(round(cas_total[8])),float_with_comma(round(cas_total_quarter[3])),float_with_comma(round(cas_total[9])),float_with_comma(round(cas_total[10])),float_with_comma(round(cas_total[11])),float_with_comma(round(cas_total_quarter[4])),float_with_comma(round(cas_total_quarter[5]))])
            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                data.append([c,float_with_comma(round(self.fn.channel_total_ly(cas,c))),float_with_comma(round(value[0])),float_with_comma(round(value[1])),float_with_comma(round(value[2])),float_with_comma(round(value[0]+value[1]+value[2])),float_with_comma(round(value[3])),float_with_comma(round(value[4])),float_with_comma(round(value[5])),float_with_comma(round(value[3]+value[4]+value[5])),float_with_comma(round(value[6])),float_with_comma(round(value[7])),float_with_comma(round(value[8])),float_with_comma(round(value[6]+value[7]+value[8])),float_with_comma(round(value[9])),float_with_comma(round(value[10])),float_with_comma(round(value[11])),float_with_comma(round(value[9]+value[10]+value[11])),float_with_comma(round(sum(value)))])

        data.append(['Total',float_with_comma(round(self.fn.all_customers_total(self.CURRENT_YEAR-1))),float_with_comma(round(self.total_for_customers[1])),float_with_comma(round(self.total_for_customers[2])),float_with_comma(round(self.total_for_customers[3])),float_with_comma(round(self.fn.all_customers_total_for_quarter(1))),float_with_comma(round(self.total_for_customers[4])),float_with_comma(round(self.total_for_customers[5])),float_with_comma(round(self.total_for_customers[6])),float_with_comma(round(self.fn.all_customers_total_for_quarter(2))),float_with_comma(round(self.total_for_customers[7])),float_with_comma(round(self.total_for_customers[8])),float_with_comma(round(self.total_for_customers[9])),float_with_comma(round(self.fn.all_customers_total_for_quarter(3))),float_with_comma(round(self.total_for_customers[10])),float_with_comma(round(self.total_for_customers[11])),float_with_comma(round(self.total_for_customers[12])),float_with_comma(round(self.fn.all_customers_total_for_quarter(4))),float_with_comma(round(self.fn.all_customers_total(self.CURRENT_YEAR)))])
        return data

    def sov_table(self):
        data = [['SOV Total']]
        data.append(['Row Labels']+self.headers)

        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            cas_last_year = self.fn.total_customer_ly(cas)

            data.append([cas,calculate_percente(cas_last_year,self.fn.all_customers_total(self.CURRENT_YEAR-1)),calculate_percente(cas_total[0],self.total_for_customers[1]),calculate_percente(cas_total[1],self.total_for_customers[2]),calculate_percente(cas_total[2],self.total_for_customers[3]),calculate_percente(cas_total_quarter[1],self.fn.all_customers_total_for_quarter(1)),calculate_percente(cas_total[3],self.total_for_customers[4]),calculate_percente(cas_total[4],self.total_for_customers[5]),calculate_percente(cas_total[5],self.total_for_customers[6]),calculate_percente(cas_total_quarter[2],self.fn.all_customers_total_for_quarter(2)),calculate_percente(cas_total[6],self.total_for_customers[7]),calculate_percente(cas_total[7],self.total_for_customers[8]),calculate_percente(cas_total[8],self.total_for_customers[9]),calculate_percente(cas_total_quarter[3],self.fn.all_customers_total_for_quarter(3)),calculate_percente(cas_total[9],self.total_for_customers[10]),calculate_percente(cas_total[10],self.total_for_customers[11]),calculate_percente(cas_total[11],self.total_for_customers[12]),calculate_percente(cas_total_quarter[4],self.fn.all_customers_total_for_quarter(4)),calculate_percente(cas_total_quarter[5],self.fn.all_customers_total(self.CURRENT_YEAR))])

            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                data.append([c,calculate_percente(self.fn.channel_total_ly(cas,c),cas_last_year),calculate_percente(value[0],cas_total[0]),calculate_percente(value[1],cas_total[1]),calculate_percente(value[2],cas_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_total_quarter[1]),calculate_percente(value[3],cas_total[3]),calculate_percente(value[4],cas_total[4]),calculate_percente(value[5],cas_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_total_quarter[2]),calculate_percente(value[6],cas_total[6]),calculate_percente(value[7],cas_total[7]),calculate_percente(value[8],cas_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_total_quarter[3]),calculate_percente(value[9],cas_total[9]),calculate_percente(value[10],cas_total[10]),calculate_percente(value[11],cas_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_total_quarter[4]),calculate_percente(sum(value),cas_total_quarter[5])])
        return data


    def sov1(self):
        sov1 = {}
        for cas in self.fn.customers:
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            cas_last_year = self.fn.total_customer_ly(cas)
            s = [float_with_comma(round(cas_last_year))]
            for i in range(1,self.QUARTER+1):
                s += [float_with_comma(round(cas_total_quarter[i]))]
            s += [float_with_comma(round(cas_total_quarter[5])), calculate_percente(cas_last_year,self.fn.all_customers_total(self.CURRENT_YEAR-1))]
            for i in range(1,self.QUARTER+1):
                s += [calculate_percente(cas_total_quarter[i],self.fn.all_customers_total_for_quarter(i))]
            s += [calculate_percente(cas_total_quarter[5],self.fn.all_customers_total(self.CURRENT_YEAR))]

            sov1["{0}_total".format(cas)] = s

            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                ch = [float_with_comma(round(self.fn.channel_total_ly(cas,c)))]
                for i in range(1,self.QUARTER+1):
                    ch += [float_with_comma(round(sum_values_for_quarter(value,i)))]
                ch += [float_with_comma(round(sum(value))),calculate_percente(self.fn.channel_total_ly(cas,c),cas_last_year)]
                for i in range(1,self.QUARTER+1):
                    ch += [calculate_percente(sum_values_for_quarter(value,i),cas_total_quarter[i])]
                ch += [calculate_percente(sum(value),cas_total_quarter[5])]
                    
                sov1["{0}_{1}".format(cas,c)] = ch

        return sov1


    def table_18_50(self):
        data = [['18-50']]
        data.append(['Row Labels']+self.headers)

        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            data.append([cas,float_with_comma(round(self.fn.total_customer_ly(cas))),float_with_comma(round(cas_total[0])),float_with_comma(round(cas_total[1])),float_with_comma(round(cas_total[2])),float_with_comma(round(cas_total_quarter[1])),float_with_comma(round(cas_total[3])),float_with_comma(round(cas_total[4])),float_with_comma(round(cas_total[5])),float_with_comma(round(cas_total_quarter[2])),float_with_comma(round(cas_total[6])),float_with_comma(round(cas_total[7])),float_with_comma(round(cas_total[8])),float_with_comma(round(cas_total_quarter[3])),float_with_comma(round(cas_total[9])),float_with_comma(round(cas_total[10])),float_with_comma(round(cas_total[11])),float_with_comma(round(cas_total_quarter[4])),float_with_comma(round(cas_total_quarter[5]))])
            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                data.append([c,float_with_comma(round(self.fn.channel_total_ly(cas,c))),float_with_comma(round(value[0])),float_with_comma(round(value[1])),float_with_comma(round(value[2])),float_with_comma(round(value[0]+value[1]+value[2])),float_with_comma(round(value[3])),float_with_comma(round(value[4])),float_with_comma(round(value[5])),float_with_comma(round(value[3]+value[4]+value[5])),float_with_comma(round(value[6])),float_with_comma(round(value[7])),float_with_comma(round(value[8])),float_with_comma(round(value[6]+value[7]+value[8])),float_with_comma(round(value[9])),float_with_comma(round(value[10])),float_with_comma(round(value[11])),float_with_comma(round(value[9]+value[10]+value[11])),float_with_comma(round(sum(value)))])

        data.append(['Total',float_with_comma(round(self.fn.all_customers_total(self.CURRENT_YEAR-1))),float_with_comma(round(self.total_for_customers[1])),float_with_comma(round(self.total_for_customers[2])),float_with_comma(round(self.total_for_customers[3])),float_with_comma(round(self.fn.all_customers_total_for_quarter(1))),float_with_comma(round(self.total_for_customers[4])),float_with_comma(round(self.total_for_customers[5])),float_with_comma(round(self.total_for_customers[6])),float_with_comma(round(self.fn.all_customers_total_for_quarter(2))),float_with_comma(round(self.total_for_customers[7])),float_with_comma(round(self.total_for_customers[8])),float_with_comma(round(self.total_for_customers[9])),float_with_comma(round(self.fn.all_customers_total_for_quarter(3))),float_with_comma(round(self.total_for_customers[10])),float_with_comma(round(self.total_for_customers[11])),float_with_comma(round(self.total_for_customers[12])),float_with_comma(round(self.fn.all_customers_total_for_quarter(4))),float_with_comma(round(self.fn.all_customers_total(self.CURRENT_YEAR)))])
        return data

    def sov_18_50(self):
        data = [['SOV 18-50']]
        data.append(['Row Labels']+self.headers)

        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            cas_last_year = self.fn.total_customer_ly(cas)

            data.append([cas,calculate_percente(cas_last_year,self.fn.all_customers_total(self.CURRENT_YEAR-1)),calculate_percente(cas_total[0],self.total_for_customers[1]),calculate_percente(cas_total[1],self.total_for_customers[2]),calculate_percente(cas_total[2],self.total_for_customers[3]),calculate_percente(cas_total_quarter[1],self.fn.all_customers_total_for_quarter(1)),calculate_percente(cas_total[3],self.total_for_customers[4]),calculate_percente(cas_total[4],self.total_for_customers[5]),calculate_percente(cas_total[5],self.total_for_customers[6]),calculate_percente(cas_total_quarter[2],self.fn.all_customers_total_for_quarter(2)),calculate_percente(cas_total[6],self.total_for_customers[7]),calculate_percente(cas_total[7],self.total_for_customers[8]),calculate_percente(cas_total[8],self.total_for_customers[9]),calculate_percente(cas_total_quarter[3],self.fn.all_customers_total_for_quarter(3)),calculate_percente(cas_total[9],self.total_for_customers[10]),calculate_percente(cas_total[10],self.total_for_customers[11]),calculate_percente(cas_total[11],self.total_for_customers[12]),calculate_percente(cas_total_quarter[4],self.fn.all_customers_total_for_quarter(4)),calculate_percente(cas_total_quarter[5],self.fn.all_customers_total(self.CURRENT_YEAR))])

            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                data.append([c,calculate_percente(self.fn.channel_total_ly(cas,c),cas_last_year),calculate_percente(value[0],cas_total[0]),calculate_percente(value[1],cas_total[1]),calculate_percente(value[2],cas_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_total_quarter[1]),calculate_percente(value[3],cas_total[3]),calculate_percente(value[4],cas_total[4]),calculate_percente(value[5],cas_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_total_quarter[2]),calculate_percente(value[6],cas_total[6]),calculate_percente(value[7],cas_total[7]),calculate_percente(value[8],cas_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_total_quarter[3]),calculate_percente(value[9],cas_total[9]),calculate_percente(value[10],cas_total[10]),calculate_percente(value[11],cas_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_total_quarter[4]),calculate_percente(sum(value),cas_total_quarter[5])])
        return data

    def sov2(self):
        sov2 = {}
        for cas in self.fn.customers:
            cas_total_quarter = self.grp_customers_quarter_v[cas]
            cas_last_year = self.fn.total_customer_ly(cas)
            s = [float_with_comma(round(cas_last_year))]
            for i in range(1,self.QUARTER+1):
                s += [float_with_comma(round(cas_total_quarter[i]))]
            s += [float_with_comma(round(cas_total_quarter[5])), calculate_percente(cas_last_year,self.fn.all_customers_total(self.CURRENT_YEAR-1))]
            for i in range(1,self.QUARTER+1):
                s += [calculate_percente(cas_total_quarter[i],self.fn.all_customers_total_for_quarter(i))]
            s += [calculate_percente(cas_total_quarter[5],self.fn.all_customers_total(self.CURRENT_YEAR))]

            sov2["{0}_total".format(cas)] = s

            for c in self.fn.channels:
                value = self.grp_channels["{0}_{1}".format(cas,c)]
                ch = [float_with_comma(round(self.fn.channel_total_ly(cas,c)))]
                for i in range(1,self.QUARTER+1):
                    ch += [float_with_comma(round(sum_values_for_quarter(value,i)))]
                ch += [float_with_comma(round(sum(value))),calculate_percente(self.fn.channel_total_ly(cas,c),cas_last_year)]
                for i in range(1,self.QUARTER+1):
                    ch += [calculate_percente(sum_values_for_quarter(value,i),cas_total_quarter[i])]
                ch += [calculate_percente(sum(value),cas_total_quarter[5])]

                sov2["{0}_{1}".format(cas,c)] = ch
        return sov2

    def shr_sov_total_table(self):
        data = [['SOV Total']]
        data.append(["SHR% of Audience TOTAL 4+ (08-24 h)"])
        data.append(['Row Labels'] + self.headers)
        v = self.fn.dm_pool_shr_total(13,31)
        data.append(['Total',v[1],v[2],v[3],v[4],v[5],v[6],v[7],v[8],v[9],v[10],v[11],v[12],v[13],v[14],v[15],v[16],v[17],v[18]])
        for d in self.fn.shr_channel_value(13,31):
            data.append(d)
        return data

    def shr_18_50_total_table(self):
        data = [['18-50']]
        data.append(["SHR% of Audience 18-50 (08-24 h)"])
        data.append(['Row Labels'] + self.headers)

        g = ('Total',)
        for i in list(range(45,63)):
            g = g + (self.fn.shr_total_value(i),)
        data.append(list(g))

        for c in self.fn.channels:
            tup = (c,)
            for i in list(range(45,63)):
                tup = tup + (self.fn.formated_shr_value(c,i),)
            data.append(list(tup))
        return data

    def kontrolni_duration(self):
        data = [['KONTROLNI ZA DURATION']]
        data.append(['DURATION'])
        data.append(['Row Labels']+self.headers[1:])

        for cas in self.fn.customers:
            cas_total = self.cas_duration[cas]

            data.append([cas,calculate_days(cas_total[0]),calculate_days(cas_total[1]),calculate_days(cas_total[2]),calculate_days(cas_total[0]+cas_total[1]+cas_total[2]),calculate_days(cas_total[3]),calculate_days(cas_total[4]),calculate_days(cas_total[5]),calculate_days(cas_total[3]+cas_total[4]+cas_total[5]),calculate_days(cas_total[6]),calculate_days(cas_total[7]),calculate_days(cas_total[8]),calculate_days(cas_total[6]+cas_total[7]+cas_total[8]),calculate_days(cas_total[9]),calculate_days(cas_total[10]),calculate_days(cas_total[11]),calculate_days(cas_total[9]+cas_total[10]+cas_total[11]),calculate_days(self.fn.sum_delta_time(cas_total))])
            for c in self.fn.channels:
                value = self.c_duration["{0}_{1}".format(cas,c)]
                data.append([c,calculate_days(value[0]),calculate_days(value[1]),calculate_days(value[2]),calculate_days(value[0]+value[1]+value[2]),calculate_days(value[3]),calculate_days(value[4]),calculate_days(value[5]),calculate_days(value[3]+value[4]+value[5]),calculate_days(value[6]),calculate_days(value[7]),calculate_days(value[8]),calculate_days(value[6]+value[7]+value[8]),calculate_days(value[9]),calculate_days(value[10]),calculate_days(value[11]),calculate_days(value[9]+value[10]+value[11]),calculate_days(self.fn.sum_delta_time(value))])

        data.append(['Total',calculate_days(self.all_duration[0]),calculate_days(self.all_duration[1]),calculate_days(self.all_duration[2]),calculate_days(self.all_duration[0]+self.all_duration[1]+self.all_duration[2]),calculate_days(self.all_duration[3]),calculate_days(self.all_duration[4]),calculate_days(self.all_duration[5]),calculate_days(self.all_duration[3]+self.all_duration[4]+self.all_duration[5]),calculate_days(self.all_duration[6]),calculate_days(self.all_duration[7]),calculate_days(self.all_duration[8]),calculate_days(self.all_duration[6]+self.all_duration[7]+self.all_duration[8]),calculate_days(self.all_duration[9]),calculate_days(self.all_duration[10]),calculate_days(self.all_duration[11]),calculate_days(self.all_duration[9]+self.all_duration[10]+self.all_duration[11]),calculate_days(self.fn.sum_delta_time(self.all_duration.values()))])
        return data

    def duration(self):
        data = [['DURATION']]
        data.append(['Row Labels']+self.headers)
        data_for_cas = {}
        for cas in self.fn.customers:
            cas_total = self.cas_duration[cas]

            data.append([cas,self.duration_ly[cas],calculate_percente(cas_total[0],self.all_duration[0]),calculate_percente(cas_total[1],self.all_duration[1]),calculate_percente(cas_total[2],self.all_duration[2]),calculate_percente(cas_total[0]+cas_total[1]+cas_total[2],self.all_duration[0]+self.all_duration[1]+self.all_duration[2]),calculate_percente(cas_total[3],self.all_duration[3]),calculate_percente(cas_total[4],self.all_duration[4]),calculate_percente(cas_total[5],self.all_duration[5]),calculate_percente(cas_total[3]+cas_total[4]+cas_total[5],self.all_duration[3]+self.all_duration[4]+self.all_duration[5]),calculate_percente(cas_total[6],self.all_duration[6]),calculate_percente(cas_total[7],self.all_duration[7]),calculate_percente(cas_total[8],self.all_duration[8]),calculate_percente(cas_total[6]+cas_total[7]+cas_total[8],self.all_duration[6]+self.all_duration[7]+self.all_duration[8]),calculate_percente(cas_total[9],self.all_duration[9]),calculate_percente(cas_total[10],self.all_duration[10]),calculate_percente(cas_total[11],self.all_duration[11]),calculate_percente(cas_total[9]+cas_total[10]+cas_total[11],self.all_duration[9]+self.all_duration[10]+self.all_duration[11]),calculate_percente(self.fn.sum_delta_time(cas_total),self.fn.sum_delta_time(self.all_duration.values()))])
            for c in self.fn.channels:
                value = self.c_duration["{0}_{1}".format(cas,c)]
                data.append([c,self.duration_ly['{0}_{1}'.format(cas,c)],calculate_percente(value[0],cas_total[0]),calculate_percente(value[1],cas_total[1]),calculate_percente(value[2],cas_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_total[0]+cas_total[1]+cas_total[2]),calculate_percente(value[3],cas_total[3]),calculate_percente(value[4],cas_total[4]),calculate_percente(value[5],cas_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_total[3]+cas_total[4]+cas_total[5]),calculate_percente(value[6],cas_total[6]),calculate_percente(value[7],cas_total[7]),calculate_percente(value[8],cas_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_total[6]+cas_total[7]+cas_total[8]),calculate_percente(value[9],cas_total[9]),calculate_percente(value[10],cas_total[10]),calculate_percente(value[11],cas_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_total[9]+cas_total[10]+cas_total[11]),calculate_percente(self.fn.sum_delta_time(value),self.fn.sum_delta_time(cas_total))])

                if c == 'Cas Media':
                    data_for_cas[cas] = (self.duration_ly['{0}_{1}'.format(cas,c)],calculate_percente(value[0],cas_total[0]),calculate_percente(value[1],cas_total[1]),calculate_percente(value[2],cas_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_total[0]+cas_total[1]+cas_total[2]),calculate_percente(value[3],cas_total[3]),calculate_percente(value[4],cas_total[4]),calculate_percente(value[5],cas_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_total[3]+cas_total[4]+cas_total[5]),calculate_percente(value[6],cas_total[6]),calculate_percente(value[7],cas_total[7]),calculate_percente(value[8],cas_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_total[6]+cas_total[7]+cas_total[8]),calculate_percente(value[9],cas_total[9]),calculate_percente(value[10],cas_total[10]),calculate_percente(value[11],cas_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_total[9]+cas_total[10]+cas_total[11]),calculate_percente(self.fn.sum_delta_time(value),self.fn.sum_delta_time(cas_total)))
        return { 'data': data, 'cas_header': data_for_cas }

    def dur(self):
        dur = {}
        for cas in self.fn.customers:
            cas_total = self.cas_duration[cas]
            t = [self.duration_ly[cas]]

            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                t += [calculate_percente(cas_total[i]+cas_total[i+1]+cas_total[i+2],self.all_duration[i]+self.all_duration[i+1]+self.all_duration[i+2])]

            t += [calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]],self.all_duration[self.fn.range_by_quarter[self.QUARTER]]),calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]+1],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]+2],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]]+cas_total[self.fn.range_by_quarter[self.QUARTER]+1]+cas_total[self.fn.range_by_quarter[self.QUARTER]+2],self.all_duration[self.fn.range_by_quarter[self.QUARTER]]+self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]+self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(self.fn.sum_delta_time(cas_total),self.fn.sum_delta_time(self.all_duration.values()))]

            dur["{0}_total".format(cas)] = t

            for c in self.fn.channels:
                value = self.c_duration["{0}_{1}".format(cas,c)]

                ch = [self.duration_ly['{0}_{1}'.format(cas,c)]]
                for q in range(1,self.QUARTER):
                    i = self.fn.range_by_quarter[q]
                    ch += [calculate_percente(value[i]+value[i+1]+value[i+2],cas_total[i]+cas_total[i+1]+cas_total[i+2])]
                    
                ch += [calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]],cas_total[self.fn.range_by_quarter[self.QUARTER]]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+1],cas_total[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+2],cas_total[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]]+value[self.fn.range_by_quarter[self.QUARTER]+1]+value[self.fn.range_by_quarter[self.QUARTER]+2],cas_total[self.fn.range_by_quarter[self.QUARTER]]+cas_total[self.fn.range_by_quarter[self.QUARTER]+1]+cas_total[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(self.fn.sum_delta_time(value),self.fn.sum_delta_time(cas_total))]
                
                dur["{0}_{1}".format(cas,c)] = ch
        return dur


    def soi_total_table(self):
        new_header = []
        for h in self.headers:
            new_header.append(h)
            new_header.append('')

        data = [['Sum of UKUPNO TOTAL 4+','SOI%']]
        data.append(['Row Labels']+new_header)

        soi_total = self.soi_total
        e2_cas_sk = self.e2_cas_sk
        all_c = self.all_c
        last_year_total = self.last_year_total

        new_dm_pool_header = [last_year_total,soi_total[0],soi_total[1],soi_total[2],soi_total[0] + soi_total[1] + soi_total[2],soi_total[3],soi_total[4],soi_total[5],soi_total[3] + soi_total[4] + soi_total[5],soi_total[6],soi_total[7],soi_total[8],soi_total[6] + soi_total[7] + soi_total[8],soi_total[9],soi_total[10],soi_total[11],soi_total[9] + soi_total[10] + soi_total[11],sum(soi_total.values())]


        dm_pool_header = [float_with_comma(round(last_year_total)),calculate_percente_round_1(last_year_total,last_year_total),float_with_comma(round(soi_total[0])),calculate_percente_round_1(soi_total[0],soi_total[0]),float_with_comma(round(soi_total[1])),calculate_percente_round_1(soi_total[1],soi_total[1]),float_with_comma(round(soi_total[2])),calculate_percente_round_1(soi_total[2],soi_total[2]),float_with_comma(round(soi_total[0] + soi_total[1] + soi_total[2])),calculate_percente_round_1(soi_total[0] + soi_total[1] + soi_total[2], soi_total[0] + soi_total[1] + soi_total[2]),float_with_comma(round(soi_total[3])),calculate_percente_round_1(soi_total[3],soi_total[3]),float_with_comma(round(soi_total[4])),calculate_percente_round_1(soi_total[4],soi_total[4]),float_with_comma(round(soi_total[5])),calculate_percente_round_1(soi_total[5],soi_total[5]),float_with_comma(round(soi_total[3] + soi_total[4] + soi_total[5])),calculate_percente_round_1(soi_total[3] + soi_total[4] + soi_total[5],soi_total[3] + soi_total[4] + soi_total[5]),float_with_comma(round(soi_total[6])),calculate_percente_round_1(soi_total[6],soi_total[6]),float_with_comma(round(soi_total[7])),calculate_percente_round_1(soi_total[7],soi_total[7]),float_with_comma(round(soi_total[8])),calculate_percente_round_1(soi_total[8],soi_total[8]),float_with_comma(round(soi_total[6] + soi_total[7] + soi_total[8])),calculate_percente_round_1(soi_total[6] + soi_total[7] + soi_total[8],soi_total[6] + soi_total[7] + soi_total[8]),float_with_comma(round(soi_total[9])),calculate_percente_round_1(soi_total[9],soi_total[9]),float_with_comma(round(soi_total[10])),calculate_percente_round_1(soi_total[10],soi_total[10]),float_with_comma(round(soi_total[11])),calculate_percente_round_1(soi_total[11],soi_total[11]),float_with_comma(round(soi_total[9] + soi_total[10] + soi_total[11])),calculate_percente_round_1(soi_total[9] + soi_total[10] + soi_total[11],soi_total[9] + soi_total[10] + soi_total[11]),float_with_comma(round(sum(soi_total.values()))),calculate_percente_round_1(sum(soi_total.values()),sum(soi_total.values()))]

        data.append(['DM Pool']+dm_pool_header)
        for c in all_c:
            value = e2_cas_sk[c]
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            elif c == 'CAS (SK, Nova Sport, Brainz)':
                last_year = self.fn.sk_nova_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)
            data.append([c,float_with_comma(round(last_year)),calculate_percente_round_1(last_year,last_year_total),float_with_comma(round(value[0])),calculate_percente_round_1(value[0],soi_total[0]),float_with_comma(round(value[1])),calculate_percente_round_1(value[1],soi_total[1]),float_with_comma(round(value[2])),calculate_percente_round_1(value[2],soi_total[2]),float_with_comma(round(value[0]+value[1]+value[2])),calculate_percente_round_1(value[0]+value[1]+value[2],soi_total[0]+soi_total[1]+soi_total[2]),float_with_comma(round(value[3])),calculate_percente_round_1(value[3],soi_total[3]),float_with_comma(round(value[4])),calculate_percente_round_1(value[4],soi_total[4]),float_with_comma(round(value[5])),calculate_percente_round_1(value[5],soi_total[5]),float_with_comma(round(value[3]+value[4]+value[5])),calculate_percente_round_1(value[3]+value[4]+value[5],soi_total[3]+soi_total[4]+soi_total[5]),float_with_comma(round(value[6])),calculate_percente_round_1(value[6],soi_total[6]),float_with_comma(round(value[7])),calculate_percente_round_1(value[7],soi_total[7]),float_with_comma(round(value[8])),calculate_percente_round_1(value[8],soi_total[8]),float_with_comma(round(value[6]+value[7]+value[8])),calculate_percente_round_1(value[6]+value[7]+value[8],soi_total[6]+soi_total[7]+soi_total[8]),float_with_comma(round(value[9])),calculate_percente_round_1(value[9],soi_total[9]),float_with_comma(round(value[10])),calculate_percente_round_1(value[10],soi_total[10]),float_with_comma(round(value[11])),calculate_percente_round_1(value[11],soi_total[11]),round(value[9]+value[10]+value[11]),calculate_percente_round_1(value[9]+value[10]+value[11],soi_total[9]+soi_total[10]+soi_total[11]),float_with_comma(round(sum(value))),calculate_percente_round_1(sum(value),sum(soi_total.values()))])
        
        # a = []
        # for index,d in enumerate(dm_pool_header):
        #     if index % 2 == 0:
        #         a += [int(d.replace(',',''))]

        return { 'data': data, 'dmp_header': dm_pool_header, 'new_dmp_header': new_dm_pool_header }
        # return { 'data': data, 'dmp_header': a, 'new_dmp_header': new_dm_pool_header }
        

    def cpp_4_plus_table(self):
        data = [['Avg CPP 4+']]
        data.append(['Row Labels']+self.headers)

        dm_pool = self.dm_pool
        # dm_total_quarter = self.dm_total_quarter
        soi_total = self.soi_total
        e2_cas_sk = self.e2_cas_sk

        data.append(['DM Pool',div(self.fn.SOI_total_ly(self.fn.e2_collection)+self.fn.cas_media_ly(),self.fn.total_customer_ly('DM Pool')),div(soi_total[0]-self.fn.sk_nova_total(1),dm_pool[0]),div(soi_total[1]-self.fn.sk_nova_total(2),dm_pool[1]),div(soi_total[2]-self.fn.sk_nova_total(3),dm_pool[2]),div(soi_total[0]+soi_total[1]+soi_total[2]-self.fn.sk_nova_total(1)-self.fn.sk_nova_total(2)-self.fn.sk_nova_total(3),dm_pool[0]+dm_pool[1]+dm_pool[2]),div(soi_total[3]-self.fn.sk_nova_total(4),dm_pool[3]),div(soi_total[4]-self.fn.sk_nova_total(5),dm_pool[4]),div(soi_total[5]-self.fn.sk_nova_total(6),dm_pool[5]),div(soi_total[3]+soi_total[4]+soi_total[5]-self.fn.sk_nova_total(4)-self.fn.sk_nova_total(5)-self.fn.sk_nova_total(6),dm_pool[3]+dm_pool[4]+dm_pool[5]),div(soi_total[6]-self.fn.sk_nova_total(7),dm_pool[6]),div(soi_total[7]-self.fn.sk_nova_total(8),dm_pool[7]),div(soi_total[8]-self.fn.sk_nova_total(9),dm_pool[8]),div(soi_total[6]+soi_total[7]+soi_total[8]-self.fn.sk_nova_total(7)-self.fn.sk_nova_total(8)-self.fn.sk_nova_total(9),dm_pool[6]+dm_pool[7]+dm_pool[8]),div(soi_total[9]-self.fn.sk_nova_total(10),dm_pool[9]),div(soi_total[10]-self.fn.sk_nova_total(11),dm_pool[10]),div(soi_total[11]-self.fn.sk_nova_total(12),dm_pool[11]),div(soi_total[9]+soi_total[10]+soi_total[11]-self.fn.sk_nova_total(10)-self.fn.sk_nova_total(11)-self.fn.sk_nova_total(12),dm_pool[9]+dm_pool[10]+dm_pool[11]),div(sum(soi_total.values())-self.fn.total_sk_nova(),sum(dm_pool.values()))])
        
        for c in self.fn.channels:
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)

            val2 = self.grp_channels["DM Pool_{0}".format(c)]
            ly1 = self.fn.channel_total_ly('DM Pool',c)
            val1 = e2_cas_sk[c]

            data.append([c,div(last_year,ly1),div(val1[0],val2[0]),div(val1[1],val2[1]),div(val1[2],val2[2]),div(val1[0]+val1[1]+val1[2],val2[0]+val2[1]+val2[2]),div(val1[3],val2[3]),div(val1[4],val2[4]),div(val1[5],val2[5]),div(val1[3]+val1[4]+val1[5],val2[3]+val2[4]+val2[5]),div(val1[6],val2[6]),div(val1[7],val2[7]),div(val1[8],val2[8]),div(val1[6]+val1[7]+val1[8],val2[6]+val2[7]+val2[8]),div(val1[9],val2[9]),div(val1[10],val2[10]),div(val1[11],val2[11]),div(val1[9]+val1[10]+val1[11],val2[9]+val2[10]+val2[11]),div(sum(val1),sum(val2))])
        return data


    def cpp_18_50_table(self):
        data = [['Avg CPP All 18-50']]
        data.append(['Row Labels']+self.headers)
        grp_channels = self.grp_channels
        dm_pool = self.dm_pool
        # dm_total_quarter = self.dm_total_quarter
        soi_total = self.soi_total
        e2_cas_sk = self.e2_cas_sk

        data.append(['DM Pool',div(self.fn.SOI_total_ly(self.fn.e2_collection)+self.fn.cas_media_ly(),self.fn.total_customer_ly('DM Pool')),div(soi_total[0]-self.fn.sk_nova_total(1),dm_pool[0]),div(soi_total[1]-self.fn.sk_nova_total(2),dm_pool[1]),div(soi_total[2]-self.fn.sk_nova_total(3),dm_pool[2]),div(soi_total[0]+soi_total[1]+soi_total[2]-self.fn.sk_nova_total(1)-self.fn.sk_nova_total(2)-self.fn.sk_nova_total(3),dm_pool[0]+dm_pool[1]+dm_pool[2]),div(soi_total[3]-self.fn.sk_nova_total(4),dm_pool[3]),div(soi_total[4]-self.fn.sk_nova_total(5),dm_pool[4]),div(soi_total[5]-self.fn.sk_nova_total(6),dm_pool[5]),div(soi_total[3]+soi_total[4]+soi_total[5]-self.fn.sk_nova_total(4)-self.fn.sk_nova_total(5)-self.fn.sk_nova_total(6),dm_pool[3]+dm_pool[4]+dm_pool[5]),div(soi_total[6]-self.fn.sk_nova_total(7),dm_pool[6]),div(soi_total[7]-self.fn.sk_nova_total(8),dm_pool[7]),div(soi_total[8]-self.fn.sk_nova_total(9),dm_pool[8]),div(soi_total[6]+soi_total[7]+soi_total[8]-self.fn.sk_nova_total(7)-self.fn.sk_nova_total(8)-self.fn.sk_nova_total(9),dm_pool[6]+dm_pool[7]+dm_pool[8]),div(soi_total[9]-self.fn.sk_nova_total(10),dm_pool[9]),div(soi_total[10]-self.fn.sk_nova_total(11),dm_pool[10]),div(soi_total[11]-self.fn.sk_nova_total(12),dm_pool[11]),div(soi_total[9]+soi_total[10]+soi_total[11]-self.fn.sk_nova_total(10)-self.fn.sk_nova_total(11)-self.fn.sk_nova_total(12),dm_pool[9]+dm_pool[10]+dm_pool[11]),div(sum(soi_total.values())-self.fn.total_sk_nova(),sum(dm_pool.values()))])

        for c in self.fn.channels:
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)

            val2 = grp_channels["DM Pool_{0}".format(c)]
            ly1 = self.fn.channel_total_ly('DM Pool',c)
            val1 = e2_cas_sk[c]
            data.append([c,div(last_year,ly1),div(val1[0],val2[0]),div(val1[1],val2[1]),div(val1[2],val2[2]),div(val1[0]+val1[1]+val1[2],val2[0]+val2[1]+val2[2]),div(val1[3],val2[3]),div(val1[4],val2[4]),div(val1[5],val2[5]),div(val1[3]+val1[4]+val1[5],val2[3]+val2[4]+val2[5]),div(val1[6],val2[6]),div(val1[7],val2[7]),div(val1[8],val2[8]),div(val1[6]+val1[7]+val1[8],val2[6]+val2[7]+val2[8]),div(val1[9],val2[9]),div(val1[10],val2[10]),div(val1[11],val2[11]),div(val1[9]+val1[10]+val1[11],val2[9]+val2[10]+val2[11]),div(sum(val1),sum(val2))])
        return data


    def ratio_total_table(self):
        data = [['POWER RATIO TOTAL 4+']]
        data.append(['Row Labels']+self.headers)

        data.append(['DM Pool','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%'])
        soi_total = self.soi_total

        for c in self.fn.channels:
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly() + self.fn.sk_nova_ly()
                value = self.fn.cas_total_for_month()
                total = sum(value.values())
            else:
                last_year = self.fn.e2_channel_total_ly(c)
                value = self.e2_cas_sk[c]
                total = sum(value)
            data.append([c,calculate_percente_2(div_2(last_year,self.last_year_total),self.fn.shr_value(c,13)),calculate_percente_2(div_2(value[0],soi_total[0]),self.fn.shr_value(c,14)), calculate_percente_2(div_2(value[1],soi_total[1]),self.fn.shr_value(c,15)), calculate_percente_2(div_2(value[2],soi_total[2]),self.fn.shr_value(c,16)), calculate_percente_2(div_2(value[0]+value[1]+value[2],soi_total[0] + soi_total[1] + soi_total[2]),self.fn.shr_value(c,17)),calculate_percente_2(div_2(value[3],soi_total[3]),self.fn.shr_value(c,18)),calculate_percente_2(div_2(value[4],soi_total[4]),self.fn.shr_value(c,19)),calculate_percente_2(div_2(value[5],soi_total[5]),self.fn.shr_value(c,20)),calculate_percente_2(div_2(value[3]+value[4]+value[5],soi_total[3] + soi_total[4] + soi_total[5]),self.fn.shr_value(c,21)),calculate_percente_2(div_2(value[6],soi_total[6]),self.fn.shr_value(c,22)),calculate_percente_2(div_2(value[7],soi_total[7]),self.fn.shr_value(c,23)),calculate_percente_2(div_2(value[8],soi_total[8]),self.fn.shr_value(c,24)),calculate_percente_2(div_2(value[6]+value[7]+value[8],soi_total[6] + soi_total[7] + soi_total[8]),self.fn.shr_value(c,25)), calculate_percente_2(div_2(value[9],soi_total[9]),self.fn.shr_value(c,26)),calculate_percente_2(div_2(value[10],soi_total[10]),self.fn.shr_value(c,27)),calculate_percente_2(div_2(value[11],soi_total[11]),self.fn.shr_value(c,28)),calculate_percente_2(div_2(value[9]+value[10]+value[11],soi_total[9] + soi_total[10] + soi_total[11]),self.fn.shr_value(c,29)),calculate_percente_2(div_2(total,sum(soi_total.values())),self.fn.shr_value(c,30))])
        return data


    def ratio_18_50_table(self):
        data = [['POWER RATIO ALL 18-50']]
        data.append(['Row Labels']+self.headers)

        data.append(['DM Pool','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%'])
        e2_cas_sk = self.e2_cas_sk
        soi_total = self.soi_total

        for c in self.fn.channels:
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly() + self.fn.sk_nova_ly()
                value = self.fn.cas_total_for_month()
                total = sum(value.values())
            else:
                last_year = self.fn.e2_channel_total_ly(c)
                value = e2_cas_sk[c]
                total = sum(value)
            data.append([c,calculate_percente_2(div_2(last_year,self.last_year_total),self.fn.shr_value(c,45)),calculate_percente_2(div_2(value[0],soi_total[0]),self.fn.shr_value(c,46)), calculate_percente_2(div_2(value[1],soi_total[1]),self.fn.shr_value(c,47)), calculate_percente_2(div_2(value[2],soi_total[2]),self.fn.shr_value(c,48)), calculate_percente_2(div_2(value[0]+value[1]+value[2],soi_total[0] + soi_total[1] + soi_total[2]),self.fn.shr_value(c,49)),calculate_percente_2(div_2(value[3],soi_total[3]),self.fn.shr_value(c,50)),calculate_percente_2(div_2(value[4],soi_total[4]),self.fn.shr_value(c,51)),calculate_percente_2(div_2(value[5],soi_total[5]),self.fn.shr_value(c,52)),calculate_percente_2(div_2(value[3]+value[4]+value[5],soi_total[3] + soi_total[4] + soi_total[5]),self.fn.shr_value(c,53)),calculate_percente_2(div_2(value[6],soi_total[6]),self.fn.shr_value(c,54)),calculate_percente_2(div_2(value[7],soi_total[7]),self.fn.shr_value(c,55)),calculate_percente_2(div_2(value[8],soi_total[8]),self.fn.shr_value(c,56)),calculate_percente_2(div_2(value[6]+value[7]+value[8],soi_total[6] + soi_total[7] + soi_total[8]),self.fn.shr_value(c,57)), calculate_percente_2(div_2(value[9],soi_total[9]),self.fn.shr_value(c,58)),calculate_percente_2(div_2(value[10],soi_total[10]),self.fn.shr_value(c,59)),calculate_percente_2(div_2(value[11],soi_total[11]),self.fn.shr_value(c,60)),calculate_percente_2(div_2(value[9]+value[10]+value[11],soi_total[9] + soi_total[10] + soi_total[11]),self.fn.shr_value(c,61)),calculate_percente_2(div_2(total,sum(soi_total.values())),self.fn.shr_value(c,62))])
        return data


    def marko_sov1_sov2_table(self,customers,sov1,sov2):
        s = ['Eq GRPs TOTAL 4+','SOV% TOTAL 4+','Eq GRPs All 18-50','SOV% All 18-50']
        new_header = []
        for x in s:
            d = [x]
            # for i in range(3+1):
            for i in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER))-1):
                d += ['']
            new_header += d
        
        data2 = [['KUPAC/TV']+new_header]

        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER),get_quarter_header(self.CURRENT_YEAR,self.QUARTER),get_quarter_header(self.CURRENT_YEAR,self.QUARTER),get_quarter_header(self.CURRENT_YEAR,self.QUARTER)]
        header_list = [item for hl in header_list for item in hl]

        data2.append(['']+header_list)

        for cas in customers:
            data2.append([cas]+sov1["{0}_total".format(cas)] + sov2["{0}_total".format(cas)])
            for c in self.fn.channels:
                data2.append([c]+sov1["{0}_{1}".format(cas,c)] + sov2["{0}_{1}".format(cas,c)])

        return data2

    def marko_shr_duration(self,customers):
        dur = self.dur()
        data = [['Row Labels', 'Shr of Duration']]
        data.append(['']+get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))

        for cas in customers:
            data.append([cas]+dur["{0}_total".format(cas)])
            for c in self.fn.channels:
                data.append([c]+dur["{0}_{1}".format(cas,c)])
        return data

    def marko_sh_total_population(self):
        data = [["Variable","SHR % TOTAL Population"]]
        data.append(['Channel\Year',self.CURRENT_YEAR-1,self.CURRENT_YEAR])
        data.append(['Months'] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))

        for c in self.fn.channels:
            tup = (c,)
            for i in self.fn.shr_total_population_range[self.QUARTER]:
                tup = tup + (self.fn.formated_shr_value(c,i),)
            data.append(list(tup))
        return data

    def marko_sh_18_50_table(self):
        data = [["Variable","SHR % ALL 18-50"]]
        data.append(['Channel\Year',self.CURRENT_YEAR-1,self.CURRENT_YEAR])
        data.append(['Months'] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))

        for c in self.fn.channels:
            tup = (c,)
            for i in self.fn.shr_total_18_50_range[self.QUARTER]:
                tup = tup + (self.fn.formated_shr_value(c,i),)
            data.append(list(tup))
        return data

    def marko_soi_table(self):
        data = [['Row Labels', 'SO%']]
        new_header = []
        for h in get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True):
            new_header.append(h)
            new_header.append('')

        data.append([''] + new_header)
        soi_total = self.soi_total

        so_by_quarter = ['DM Pool',float_with_comma(round(self.last_year_total)),calculate_percente_round_1(self.last_year_total,self.last_year_total)]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            so_by_quarter += [float_with_comma(round(soi_total[i] + soi_total[i+1] + soi_total[i+2])),calculate_percente_round_1(soi_total[i] + soi_total[i+1] + soi_total[i+2], soi_total[i] + soi_total[i+1] + soi_total[i+2])]

        so_by_quarter += [float_with_comma(round(soi_total[self.fn.range_by_quarter[self.QUARTER]])),calculate_percente_round_1(soi_total[self.fn.range_by_quarter[self.QUARTER]],soi_total[self.fn.range_by_quarter[self.QUARTER]]),float_with_comma(round(soi_total[self.fn.range_by_quarter[self.QUARTER]+1])),calculate_percente_round_1(soi_total[self.fn.range_by_quarter[self.QUARTER]+1],soi_total[self.fn.range_by_quarter[self.QUARTER]+1]),float_with_comma(round(soi_total[self.fn.range_by_quarter[self.QUARTER]+2])),calculate_percente_round_1(soi_total[self.fn.range_by_quarter[self.QUARTER]+2],soi_total[self.fn.range_by_quarter[self.QUARTER]+2])]
    
        i = self.fn.range_by_quarter[self.QUARTER]
        so_by_quarter += [float_with_comma(round(soi_total[i] + soi_total[i+1] + soi_total[i+2])),calculate_percente_round_1(soi_total[i] + soi_total[i+1] + soi_total[i+2], soi_total[i] + soi_total[i+1] + soi_total[i+2])]

        so_by_quarter += [float_with_comma(round(sum(soi_total.values()))),calculate_percente_round_1(sum(soi_total.values()),sum(soi_total.values()))]

        data.append(so_by_quarter)

        for c in self.all_c:
            a = []
            value = self.e2_cas_sk[c]
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            elif c == 'CAS (SK, Nova Sport, Brainz)':
                last_year = self.fn.sk_nova_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)

            a += [c, float_with_comma(round(last_year)),calculate_percente_round_1(last_year,self.last_year_total)]
            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                a += [float_with_comma(round(value[i]+value[i+1]+value[i+2])),calculate_percente_round_1(value[i]+value[i+1]+value[i+2],soi_total[i]+soi_total[i+1]+soi_total[i+2])]

            a += [float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]],soi_total[self.fn.range_by_quarter[self.QUARTER]]),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+1])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]+1],soi_total[self.fn.range_by_quarter[self.QUARTER]+1]),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+2])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]+2],soi_total[self.fn.range_by_quarter[self.QUARTER]+2])]

            i2 = self.fn.range_by_quarter[self.QUARTER]
            a += [float_with_comma(round(value[i2]+value[i2+1]+value[i2+2])),calculate_percente_round_1(value[i2]+value[i2+1]+value[i2+2],soi_total[i2]+soi_total[i2+1]+soi_total[i2+2])]

            a += [float_with_comma(round(sum(value))),calculate_percente_round_1(sum(value),sum(soi_total.values()))]
            data.append(a)
        return data


    def marko_cpp_4_plus(self):
        data = [['Row Labels', 'Avg CPP 4+']]
        data.append([''] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))
        soi_total = self.soi_total
        # dm_total_quarter = self.dm_total_quarter
        dm_pool = self.dm_pool
        e2_cas_sk = self.e2_cas_sk

        cpp_4plus_by_quarter = ['DM Pool',div(self.fn.SOI_total_ly(self.fn.e2_collection)+self.fn.cas_media_ly(),self.fn.total_customer_ly('DM Pool'))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            cpp_4plus_by_quarter += [div(soi_total[i]+soi_total[i+1]+soi_total[i+2]-self.fn.sk_nova_total(i+1)-self.fn.sk_nova_total(i+2)-self.fn.sk_nova_total(i+3),dm_pool[i]+dm_pool[i+1]+dm_pool[i+2])]

        cpp_4plus_by_quarter += [div(soi_total[self.fn.range_by_quarter[self.QUARTER]]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+1),dm_pool[self.fn.range_by_quarter[self.QUARTER]]),div(soi_total[self.fn.range_by_quarter[self.QUARTER]+1]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+2),dm_pool[self.fn.range_by_quarter[self.QUARTER]+1]),div(soi_total[self.fn.range_by_quarter[self.QUARTER]+2]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+3),dm_pool[self.fn.range_by_quarter[self.QUARTER]+2])]

        i = self.fn.range_by_quarter[self.QUARTER]
        cpp_4plus_by_quarter += [div(soi_total[i]+soi_total[i+1]+soi_total[i+2]-self.fn.sk_nova_total(i+1)-self.fn.sk_nova_total(i+2)-self.fn.sk_nova_total(i+3),dm_pool[i]+dm_pool[i+1]+dm_pool[i+2])]

        cpp_4plus_by_quarter += [div(sum(soi_total.values())-self.fn.total_sk_nova(),sum(dm_pool.values()))]

        data.append(cpp_4plus_by_quarter)

        for c in self.fn.channels:
            a = []
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)
            val2 = self.fn.find_by_channel('DM Pool',c)
            ly1 = self.fn.channel_total_ly('DM Pool',c)
            val1 = e2_cas_sk[c]

            a += [c,div(last_year,ly1)]
            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                a += [div(val1[i]+val1[i+1]+val1[i+2],val2[i]+val2[i+1]+val2[i+2])]

            a += [div(val1[self.fn.range_by_quarter[self.QUARTER]],val2[self.fn.range_by_quarter[self.QUARTER]]), div(val1[self.fn.range_by_quarter[self.QUARTER]+1],val2[self.fn.range_by_quarter[self.QUARTER]+1]),div(val1[self.fn.range_by_quarter[self.QUARTER]+2],val2[self.fn.range_by_quarter[self.QUARTER]+2])]

            i2 = self.fn.range_by_quarter[self.QUARTER]
            a += [div(val1[i2]+val1[i2+1]+val1[i2+2],val2[i2]+val2[i2+1]+val2[i2+2])]

            a += [div(sum(val1),sum(val2))]

            data.append(a)
        return data

    
    def marko_cpp_18_50(self):
        data = [['Row Labels', 'Avg CPP ALL 18-50']]
        data.append([''] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))
        soi_total = self.soi_total
        dm_pool = self.dm_pool
        # dm_total_quarter = self.dm_total_quarter
        e2_cas_sk = self.e2_cas_sk

        cpp_18_50_by_quarter = ['DM Pool',div(self.fn.SOI_total_ly(self.fn.e2_collection)+self.fn.cas_media_ly(),self.fn.total_customer_ly('DM Pool'))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q] 
            cpp_18_50_by_quarter += [div(soi_total[i]+soi_total[i+1]+soi_total[i+2]-self.fn.sk_nova_total(i+1)-self.fn.sk_nova_total(i+2)-self.fn.sk_nova_total(i+3),dm_pool[i]+dm_pool[i+1]+dm_pool[i+2])]

        cpp_18_50_by_quarter += [div(soi_total[self.fn.range_by_quarter[self.QUARTER]]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+1),dm_pool[self.fn.range_by_quarter[self.QUARTER]]),div(soi_total[self.fn.range_by_quarter[self.QUARTER]+1]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+2),dm_pool[self.fn.range_by_quarter[self.QUARTER]+1]),div(soi_total[self.fn.range_by_quarter[self.QUARTER]+2]-self.fn.sk_nova_total(self.fn.range_by_quarter[self.QUARTER]+3),dm_pool[self.fn.range_by_quarter[self.QUARTER]+2])]

        i = self.fn.range_by_quarter[self.QUARTER]
        cpp_18_50_by_quarter += [div(soi_total[i]+soi_total[i+1]+soi_total[i+2]-self.fn.sk_nova_total(i+1)-self.fn.sk_nova_total(i+2)-self.fn.sk_nova_total(i+3),dm_pool[i]+dm_pool[i+1]+dm_pool[i+2])]

        cpp_18_50_by_quarter += [div(sum(soi_total.values())-self.fn.total_sk_nova(),sum(dm_pool.values()))]

        data.append(cpp_18_50_by_quarter)

        for c in self.fn.channels:
            a = []
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly()
            else:
                last_year = self.fn.e2_channel_total_ly(c)
            val2 = self.fn.find_by_channel('DM Pool',c)
            ly1 = self.fn.channel_total_ly('DM Pool',c)
            val1 = e2_cas_sk[c]

            a += [c,div(last_year,ly1)]
            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                a += [div(val1[i]+val1[i+1]+val1[i+2],val2[i]+val2[i+1]+val2[i+2])]

            a += [div(val1[self.fn.range_by_quarter[self.QUARTER]],val2[self.fn.range_by_quarter[self.QUARTER]]), div(val1[self.fn.range_by_quarter[self.QUARTER]+1],val2[self.fn.range_by_quarter[self.QUARTER]+1]),div(val1[self.fn.range_by_quarter[self.QUARTER]+2],val2[self.fn.range_by_quarter[self.QUARTER]+2])]

            i2 = self.fn.range_by_quarter[self.QUARTER]
            a += [div(val1[i2]+val1[i2+1]+val1[i2+2],val2[i2]+val2[i2+1]+val2[i2+2])]

            a += [div(sum(val1),sum(val2))]

            data.append(a)
        return data


    def marko_ratio_4_plus(self):
        data = [['Row Labels', 'POWER RATIO TOTAL 4+']]
        data.append([''] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))

        header_ratio = ['DM Pool']
        for g in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))):
            header_ratio += ['100%']

        data.append(header_ratio)

        shr_4_plus_by_quarter = {
            1: 17,
            2: 21,
            3: 25,
            4: 29
        }

        e2_cas_sk = self.e2_cas_sk
        soi_total = self.soi_total

        for c in self.fn.channels:
            a = []
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly() + self.fn.sk_nova_ly()
                value = self.fn.cas_total_for_month()
                total = sum(value.values())
            else:
                last_year = self.fn.e2_channel_total_ly(c)
                value = e2_cas_sk[c]
                total = sum(value)

            a += [c,calculate_percente_2(div_2(last_year,self.last_year_total),self.fn.shr_value(c,13))]
            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                shr = shr_4_plus_by_quarter[q]
                a += [calculate_percente_2(div_2(value[i]+value[i+1]+value[i+2],soi_total[i] + soi_total[i+1] + soi_total[i+2]),self.fn.shr_value(c,shr))]

            a += [calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]],soi_total[self.fn.range_by_quarter[self.QUARTER]]),self.fn.shr_value(c,shr_4_plus_by_quarter[self.QUARTER]-3)),calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]+1],soi_total[self.fn.range_by_quarter[self.QUARTER]+1]),self.fn.shr_value(c,shr_4_plus_by_quarter[self.QUARTER]-2)),calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]+2],soi_total[self.fn.range_by_quarter[self.QUARTER]+2]),self.fn.shr_value(c,shr_4_plus_by_quarter[self.QUARTER]-1))]

            i2 = self.fn.range_by_quarter[self.QUARTER]
            shr2 = shr_4_plus_by_quarter[self.QUARTER]
            a += [calculate_percente_2(div_2(value[i2]+value[i2+1]+value[i2+2],soi_total[i2] + soi_total[i2+1] + soi_total[i2+2]),self.fn.shr_value(c,shr2))]

            a += [calculate_percente_2(div_2(total,sum(soi_total.values())),self.fn.shr_value(c,30))]

            data.append(a)
        return data

    

    def marko_ratio_18_50(self):
        data = [['Row Labels','POWER RATIO ALL 18-50']]

        data.append([''] + get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))

        header_ratio = ['DM Pool']
        for g in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))):
            header_ratio += ['100%']

        data.append(header_ratio)

        shr_18_50_by_quarter = {
            1: 49,
            2: 53,
            3: 57,
            4: 61
        }

        soi_total = self.soi_total
        e2_cas_sk = self.e2_cas_sk

        for c in self.fn.channels:
            a = []
            if c == 'Cas Media':
                last_year = self.fn.cas_media_ly() + self.fn.sk_nova_ly()
                value = self.fn.cas_total_for_month()
                total = sum(value.values())
            else:
                last_year = self.fn.e2_channel_total_ly(c)
                value = e2_cas_sk[c]
                total = sum(value)

            a += [c,calculate_percente_2(div_2(last_year,self.last_year_total),self.fn.shr_value(c,45))]

            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                shr = shr_18_50_by_quarter[q]
                a += [calculate_percente_2(div_2(value[i]+value[i+1]+value[i+2],soi_total[i] + soi_total[i+1] + soi_total[i+2]),self.fn.shr_value(c,shr))]


            a += [calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]],soi_total[self.fn.range_by_quarter[self.QUARTER]]),self.fn.shr_value(c,shr_18_50_by_quarter[self.QUARTER]-3)),calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]+1],soi_total[self.fn.range_by_quarter[self.QUARTER]+1]),self.fn.shr_value(c,shr_18_50_by_quarter[self.QUARTER]-2)),calculate_percente_2(div_2(value[self.fn.range_by_quarter[self.QUARTER]+2],soi_total[self.fn.range_by_quarter[self.QUARTER]+2]),self.fn.shr_value(c,shr_18_50_by_quarter[self.QUARTER]-1))]

            i2 = self.fn.range_by_quarter[self.QUARTER]
            shr2 = shr_18_50_by_quarter[self.QUARTER]
            a += [calculate_percente_2(div_2(value[i2]+value[i2+1]+value[i2+2],soi_total[i2] + soi_total[i2+1] + soi_total[i2+2]),self.fn.shr_value(c,shr2))]

            a += [calculate_percente_2(div_2(total,sum(soi_total.values())),self.fn.shr_value(c,62))]

            data.append(a)
        return data
