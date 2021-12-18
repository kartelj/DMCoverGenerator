# -*- coding: utf-8 -*- 

from formulas.helpers.numbers import *
import formulas.cas_functions as cf
import style.cover as style

class CoverCas():

    def __init__(self,year,quarter,sources,collection=None):
        self.sources = sources
        self.collection = collection
        self.CURRENT_YEAR = year
        self.QUARTER = quarter

        if collection != None:
            self.fn = cf.CasFunctions(year,quarter,collection)
            if 'grp' in sources:
                self.grp_channels = self.fn.channels_for_customer2()
                self.grp_customers_v = self.fn.total_customer_for_month(self.grp_channels)
                self.total_for_customers = self.fn.total_amount_for_months(self.grp_customers_v)
                self.n1_nova_s_total_for_months = self.fn.n1_nova_s_total_for_months(self.grp_channels)
                self.cas_others_total_for_months = self.fn.cas_others_total_for_months(self.grp_channels)
                self.total_customer_ly = self.fn.total_customer_ly()
                self.cas_others_total_customer_ly = self.fn.cas_others_total_customer_ly()
            if 'duration' in sources:
                self.c_duration = self.fn.duration_channel_total()
                self.cas_duration = self.fn.cas_duration(self.c_duration)
                self.all_duration = self.fn.total_duration_for_months(self.cas_duration)
                self.n1_nova_s_duration = self.fn.n1_nova_s_duration(self.c_duration)
                self.cas_others_duration = self.fn.cas_others_duration(self.c_duration)
                self.duration_total_by_month = self.fn.duration_total_by_month()
                self.duration_ly = self.fn.cas_duration_ly
            if 'soi' in sources:
                self.soi_total2 = self.fn.e2_channel_total()
                self.cas_media_total = self.fn.cas_media_total(self.soi_total2)
                self.cas_media_agb = self.fn.drugi_red_total(self.soi_total2)
                self.soi_cas_others = self.fn.soi_cas_others(self.soi_total2)
                self.cas_media_channels_agb = self.fn.cas_media_channels_agb(self.soi_total2)
                self.soi_agb_channal = self.fn.soi_agb_channal(self.soi_total2)

                self.soi_channel_total_ly = self.fn.soi_channel_total_ly()
                self.cas_media_agb_ly = self.fn.cas_media_agb_ly(self.soi_channel_total_ly)
                self.soi_cas_others_channels_ly = self.fn.soi_cas_others_channels_ly(self.soi_channel_total_ly)
                self.cas_media_channels_ly = self.fn.cas_media_channels_ly(self.soi_channel_total_ly)

        self.headers = ["{0} FY".format(self.CURRENT_YEAR-1), "{0} JAN".format(self.CURRENT_YEAR), "{0} FEB".format(self.CURRENT_YEAR), "{0} MAR".format(self.CURRENT_YEAR), "{0} Q1".format(self.CURRENT_YEAR), "{0} APR".format(self.CURRENT_YEAR), "{0} MAJ".format(self.CURRENT_YEAR), "{0} JUN".format(self.CURRENT_YEAR), "{0} Q2".format(self.CURRENT_YEAR), "{0} JUL".format(self.CURRENT_YEAR), "{0} AVG".format(self.CURRENT_YEAR), "{0} SEP".format(self.CURRENT_YEAR), "{0} Q3".format(self.CURRENT_YEAR), "{0} OKT".format(self.CURRENT_YEAR), "{0} NOV".format(self.CURRENT_YEAR), "{0} DEC".format(self.CURRENT_YEAR), "{0} Q4".format(self.CURRENT_YEAR), "{0} ytd".format(self.CURRENT_YEAR)]



    def total_table(self):
        data = [['Total']]
        data.append(['Row Labels']+self.headers)

        n1 = self.n1_nova_s_total_for_months['N1']
        nova_s = self.n1_nova_s_total_for_months['Nova S']

        data.append(['Grand Total',float_with_comma(round(sum(self.total_customer_ly.values()))),float_with_comma(round(self.total_for_customers[1])),float_with_comma(round(self.total_for_customers[2])),float_with_comma(round(self.total_for_customers[3])),float_with_comma(round(self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3])),float_with_comma(round(self.total_for_customers[4])),float_with_comma(round(self.total_for_customers[5])),float_with_comma(round(self.total_for_customers[6])),float_with_comma(round(self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6])),float_with_comma(round(self.total_for_customers[7])),float_with_comma(round(self.total_for_customers[8])),float_with_comma(round(self.total_for_customers[9])),float_with_comma(round(self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9])),float_with_comma(round(self.total_for_customers[10])),float_with_comma(round(self.total_for_customers[11])),float_with_comma(round(self.total_for_customers[12])),float_with_comma(round(self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12])),float_with_comma(round(sum(self.total_for_customers.values())))])

        data.append(['Nova S',float_with_comma(round(self.fn.channel_total_ly_no_customer('Nova S'))),float_with_comma(round(nova_s[0])),float_with_comma(round(nova_s[1])),float_with_comma(round(nova_s[2])),float_with_comma(round(nova_s[0]+nova_s[1]+nova_s[2])),float_with_comma(round(nova_s[3])),float_with_comma(round(nova_s[4])),float_with_comma(round(nova_s[5])),float_with_comma(round(nova_s[3]+nova_s[4]+nova_s[5])),float_with_comma(round(nova_s[6])),float_with_comma(round(nova_s[7])),float_with_comma(round(nova_s[8])),float_with_comma(round(nova_s[6]+nova_s[7]+nova_s[8])),float_with_comma(round(nova_s[9])),float_with_comma(round(nova_s[10])),float_with_comma(round(nova_s[11])),float_with_comma(round(nova_s[9]+nova_s[10]+nova_s[11])),float_with_comma(round(sum(nova_s.values())))])

        data.append(['N1',float_with_comma(round(self.fn.channel_total_ly_no_customer('N1'))),float_with_comma(round(n1[0])),float_with_comma(round(n1[1])),float_with_comma(round(n1[2])),float_with_comma(round(n1[0]+n1[1]+n1[2])),float_with_comma(round(n1[3])),float_with_comma(round(n1[4])),float_with_comma(round(n1[5])),float_with_comma(round(n1[3]+n1[4]+n1[5])),float_with_comma(round(n1[6])),float_with_comma(round(n1[7])),float_with_comma(round(n1[8])),float_with_comma(round(n1[6]+n1[7]+n1[8])),float_with_comma(round(n1[9])),float_with_comma(round(n1[10])),float_with_comma(round(n1[11])),float_with_comma(round(n1[9]+n1[10]+n1[11])),float_with_comma(round(sum(n1.values())))])


        data.append(['CAS Others (w/o SK, Nova Sport, Brainz)',float_with_comma(round(sum(self.total_customer_ly.values())-self.fn.channel_total_ly_no_customer('Nova S')-self.fn.channel_total_ly_no_customer('N1'))),float_with_comma(round(self.total_for_customers[1]-nova_s[0]-n1[0])),float_with_comma(round(self.total_for_customers[2]-nova_s[1]-n1[1])),float_with_comma(round(self.total_for_customers[3]-nova_s[2]-n1[2])),
        float_with_comma(round(self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]-nova_s[0]-nova_s[1]-nova_s[2]-n1[0]-n1[1]-n1[2])),float_with_comma(round(self.total_for_customers[4]-nova_s[3]-n1[3])),
        float_with_comma(round(self.total_for_customers[5]-nova_s[4]-n1[4])),float_with_comma(round(self.total_for_customers[6]-nova_s[5]-n1[5])),float_with_comma(round(self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]-nova_s[3]-nova_s[4]-nova_s[5]-n1[3]-n1[4]-n1[5])),float_with_comma(round(self.total_for_customers[7]-nova_s[6]-n1[6])),float_with_comma(round(self.total_for_customers[8]-nova_s[7]-n1[7])),float_with_comma(round(self.total_for_customers[9]-nova_s[8]-n1[8])),float_with_comma(round(self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]-nova_s[6]-nova_s[7]-nova_s[8]-n1[6]-n1[7]-n1[8])),float_with_comma(round(self.total_for_customers[10]-nova_s[9]-n1[9])),float_with_comma(round(self.total_for_customers[11]-nova_s[10]-n1[10])),float_with_comma(round(self.total_for_customers[12]-nova_s[11]-n1[11])),float_with_comma(round(self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]-nova_s[9]-nova_s[10]-nova_s[11]-n1[9]-n1[10]-n1[11])),float_with_comma(round(sum(self.total_for_customers.values())-sum(n1.values())-sum(nova_s.values())))])

        channels = self.fn.cas_other_channels
        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            if cas == 'DM Pool':
                data.append(['CAS Media',float_with_comma(round(self.total_customer_ly[cas])),float_with_comma(round(cas_total[0])),float_with_comma(round(cas_total[1])),float_with_comma(round(cas_total[2])),float_with_comma(round(cas_total[0]+cas_total[1]+cas_total[2])),float_with_comma(round(cas_total[3])),float_with_comma(round(cas_total[4])),float_with_comma(round(cas_total[5])),float_with_comma(round(cas_total[3]+cas_total[4]+cas_total[5])),float_with_comma(round(cas_total[6])),float_with_comma(round(cas_total[7])),float_with_comma(round(cas_total[8])),float_with_comma(round(cas_total[6]+cas_total[7]+cas_total[8])),float_with_comma(round(cas_total[9])),float_with_comma(round(cas_total[10])),float_with_comma(round(cas_total[11])),float_with_comma(round(cas_total[9]+cas_total[10]+cas_total[11])),float_with_comma(round(sum(cas_total.values())))])
                data.append(['CAS Media AGB',float_with_comma(round(self.total_customer_ly[cas])),float_with_comma(round(cas_total[0])),float_with_comma(round(cas_total[1])),float_with_comma(round(cas_total[2])),float_with_comma(round(cas_total[0]+cas_total[1]+cas_total[2])),float_with_comma(round(cas_total[3])),float_with_comma(round(cas_total[4])),float_with_comma(round(cas_total[5])),float_with_comma(round(cas_total[3]+cas_total[4]+cas_total[5])),float_with_comma(round(cas_total[6])),float_with_comma(round(cas_total[7])),float_with_comma(round(cas_total[8])),float_with_comma(round(cas_total[6]+cas_total[7]+cas_total[8])),float_with_comma(round(cas_total[9])),float_with_comma(round(cas_total[10])),float_with_comma(round(cas_total[11])),float_with_comma(round(cas_total[9]+cas_total[10]+cas_total[11])),float_with_comma(round(sum(cas_total.values())))])

            data.append([cas,float_with_comma(round(self.total_customer_ly[cas])),float_with_comma(round(cas_total[0])),float_with_comma(round(cas_total[1])),float_with_comma(round(cas_total[2])),float_with_comma(round(cas_total[0]+cas_total[1]+cas_total[2])),float_with_comma(round(cas_total[3])),float_with_comma(round(cas_total[4])),float_with_comma(round(cas_total[5])),float_with_comma(round(cas_total[3]+cas_total[4]+cas_total[5])),float_with_comma(round(cas_total[6])),float_with_comma(round(cas_total[7])),float_with_comma(round(cas_total[8])),float_with_comma(round(cas_total[6]+cas_total[7]+cas_total[8])),float_with_comma(round(cas_total[9])),float_with_comma(round(cas_total[10])),float_with_comma(round(cas_total[11])),float_with_comma(round(cas_total[9]+cas_total[10]+cas_total[11])),float_with_comma(round(sum(cas_total.values())))])
            for c in channels:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    last_year = self.cas_others_total_customer_ly[cas]
                    value = self.cas_others_total_for_months[cas]
                    total_amount = sum(value.values())
                else:
                    value = self.grp_channels["{0}_{1}".format(cas,c)]
                    last_year = self.fn.channel_total_ly(cas,c)
                    total_amount = sum(value)

                data.append([c,float_with_comma(round(last_year)),float_with_comma(round(value[0])),float_with_comma(round(value[1])),float_with_comma(round(value[2])),float_with_comma(round(value[0]+value[1]+value[2])),float_with_comma(round(value[3])),float_with_comma(round(value[4])),float_with_comma(round(value[5])),float_with_comma(round(value[3]+value[4]+value[5])),float_with_comma(round(value[6])),float_with_comma(round(value[7])),float_with_comma(round(value[8])),float_with_comma(round(value[6]+value[7]+value[8])),float_with_comma(round(value[9])),float_with_comma(round(value[10])),float_with_comma(round(value[11])),float_with_comma(round(value[9]+value[10]+value[11])),float_with_comma(round(total_amount))])

        data.append(['Total',float_with_comma(round(sum(self.total_customer_ly.values()))),float_with_comma(round(self.total_for_customers[1])),float_with_comma(round(self.total_for_customers[2])),float_with_comma(round(self.total_for_customers[3])),float_with_comma(round(self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3])),float_with_comma(round(self.total_for_customers[4])),float_with_comma(round(self.total_for_customers[5])),float_with_comma(round(self.total_for_customers[6])),float_with_comma(round(self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6])),float_with_comma(round(self.total_for_customers[7])),float_with_comma(round(self.total_for_customers[8])),float_with_comma(round(self.total_for_customers[9])),float_with_comma(round(self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9])),float_with_comma(round(self.total_for_customers[10])),float_with_comma(round(self.total_for_customers[11])),float_with_comma(round(self.total_for_customers[12])),float_with_comma(round(self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12])),float_with_comma(round(sum(self.total_for_customers.values())))])
        return data



    def sov_table(self):
        data = [['SOV Total']]
        data.append(['Row Labels']+self.headers)

        n1 = self.n1_nova_s_total_for_months['N1']
        nova_s = self.n1_nova_s_total_for_months['Nova S']

        data.append(['Grand Total',calculate_percente(sum(self.total_customer_ly.values()),sum(self.total_customer_ly.values())),calculate_percente(self.total_for_customers[1],self.total_for_customers[1]),calculate_percente(self.total_for_customers[2],self.total_for_customers[2]),calculate_percente(self.total_for_customers[3],self.total_for_customers[3]),calculate_percente(self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3],self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]),calculate_percente(self.total_for_customers[4],self.total_for_customers[4]),calculate_percente(self.total_for_customers[5],self.total_for_customers[5]),calculate_percente(self.total_for_customers[6],self.total_for_customers[6]),calculate_percente(self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6],self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]),calculate_percente(self.total_for_customers[7],self.total_for_customers[7]),calculate_percente(self.total_for_customers[8],self.total_for_customers[8]),calculate_percente(self.total_for_customers[9],self.total_for_customers[9]),calculate_percente(self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9],self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]),calculate_percente(self.total_for_customers[10],self.total_for_customers[10]),calculate_percente(self.total_for_customers[11],self.total_for_customers[11]),calculate_percente(self.total_for_customers[12],self.total_for_customers[12]),calculate_percente(self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12],self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]),calculate_percente(sum(self.total_for_customers.values()),sum(self.total_for_customers.values()))])

        data.append(['Nova S',calculate_percente(self.fn.channel_total_ly_no_customer('Nova S'),sum(self.total_customer_ly.values())),calculate_percente(nova_s[0],self.total_for_customers[1]),calculate_percente(nova_s[1],self.total_for_customers[2]),calculate_percente(nova_s[2],self.total_for_customers[3]),calculate_percente(nova_s[0]+nova_s[1]+nova_s[2],self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]),calculate_percente(nova_s[3],self.total_for_customers[4]),calculate_percente(nova_s[4],self.total_for_customers[5]),calculate_percente(nova_s[5],self.total_for_customers[6]),calculate_percente(nova_s[3]+nova_s[4]+nova_s[5],self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]),calculate_percente(nova_s[6],self.total_for_customers[7]),calculate_percente(nova_s[7],self.total_for_customers[8]),calculate_percente(nova_s[8],self.total_for_customers[9]),calculate_percente(nova_s[6]+nova_s[7]+nova_s[8],self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]),calculate_percente(nova_s[9],self.total_for_customers[10]),calculate_percente(nova_s[10],self.total_for_customers[11]),calculate_percente(nova_s[11],self.total_for_customers[12]),calculate_percente(nova_s[9]+nova_s[10]+nova_s[11],self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]),calculate_percente(sum(nova_s.values()),sum(self.total_for_customers.values()))])

        data.append(['N1',calculate_percente(self.fn.channel_total_ly_no_customer('Nova S'),sum(self.total_customer_ly.values())),calculate_percente(n1[0],self.total_for_customers[1]),calculate_percente(n1[1],self.total_for_customers[2]),calculate_percente(n1[2],self.total_for_customers[3]),calculate_percente(n1[0]+n1[1]+n1[2],self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]),calculate_percente(n1[3],self.total_for_customers[4]),calculate_percente(n1[4],self.total_for_customers[5]),calculate_percente(n1[5],self.total_for_customers[6]),calculate_percente(n1[3]+n1[4]+n1[5],self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]),calculate_percente(n1[6],self.total_for_customers[7]),calculate_percente(n1[7],self.total_for_customers[8]),calculate_percente(n1[8],self.total_for_customers[9]),calculate_percente(n1[6]+n1[7]+n1[8],self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]),calculate_percente(n1[9],self.total_for_customers[10]),calculate_percente(n1[10],self.total_for_customers[11]),calculate_percente(n1[11],self.total_for_customers[12]),calculate_percente(n1[9]+n1[10]+n1[11],self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]),calculate_percente(sum(n1.values()),sum(self.total_for_customers.values()))])

        data.append(['CAS Others (w/o SK, Nova Sport, Brainz)',calculate_percente(sum(self.total_customer_ly.values())-self.fn.channel_total_ly_no_customer('Nova S')-self.fn.channel_total_ly_no_customer('N1'),sum(self.total_customer_ly.values())),calculate_percente(self.total_for_customers[1]-nova_s[0]-n1[0],self.total_for_customers[1]),calculate_percente(self.total_for_customers[2]-nova_s[1]-n1[1],self.total_for_customers[2]),calculate_percente(self.total_for_customers[3]-nova_s[2]-n1[2],self.total_for_customers[3]),calculate_percente(self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]-nova_s[0]-nova_s[1]-nova_s[2]-n1[0]-n1[1]-n1[2],self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]),calculate_percente(self.total_for_customers[4]-nova_s[3]-n1[3],self.total_for_customers[4]),calculate_percente(self.total_for_customers[5]-nova_s[4]-n1[4],self.total_for_customers[5]),calculate_percente(self.total_for_customers[6]-nova_s[5]-n1[5],self.total_for_customers[6]),calculate_percente(self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]-nova_s[3]-nova_s[4]-nova_s[5]-n1[3]-n1[4]-n1[5],self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]),calculate_percente(self.total_for_customers[7]-nova_s[6]-n1[6],self.total_for_customers[7]),calculate_percente(self.total_for_customers[8]-nova_s[7]-n1[7],self.total_for_customers[8]),calculate_percente(self.total_for_customers[9]-nova_s[8]-n1[8],self.total_for_customers[9]),calculate_percente(self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]-nova_s[6]-nova_s[7]-nova_s[8]-n1[6]-n1[7]-n1[8],self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]),calculate_percente(self.total_for_customers[10]-nova_s[9]-n1[9],self.total_for_customers[10]),calculate_percente(self.total_for_customers[11]-nova_s[10]-n1[10],self.total_for_customers[11]),calculate_percente(self.total_for_customers[12]-nova_s[11]-n1[11],self.total_for_customers[12]),calculate_percente(n1[9]+n1[10]+n1[11],self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]),calculate_percente(sum(self.total_for_customers.values())-sum(n1.values())-sum(nova_s.values()),sum(self.total_for_customers.values()))])

        channels = self.fn.cas_other_channels
        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            if cas == 'DM Pool':
                data.append(['CAS Media','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%'])
                data.append(['CAS Media AGB','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%','100%'])

            data.append([cas,calculate_percente(self.total_customer_ly[cas],sum(self.total_customer_ly.values())),calculate_percente(cas_total[0],self.total_for_customers[1]),calculate_percente(cas_total[1],self.total_for_customers[2]),calculate_percente(cas_total[2],self.total_for_customers[3]),calculate_percente(cas_total[0]+cas_total[1]+cas_total[2],self.total_for_customers[1]+self.total_for_customers[2]+self.total_for_customers[3]),calculate_percente(cas_total[3],self.total_for_customers[4]),calculate_percente(cas_total[4],self.total_for_customers[5]),calculate_percente(cas_total[5],self.total_for_customers[6]),calculate_percente(cas_total[3]+cas_total[4]+cas_total[5],self.total_for_customers[4]+self.total_for_customers[5]+self.total_for_customers[6]),calculate_percente(cas_total[6],self.total_for_customers[7]),calculate_percente(cas_total[7],self.total_for_customers[8]),calculate_percente(cas_total[8],self.total_for_customers[9]),calculate_percente(cas_total[6]+cas_total[7]+cas_total[8],self.total_for_customers[7]+self.total_for_customers[8]+self.total_for_customers[9]),calculate_percente(cas_total[9],self.total_for_customers[10]),calculate_percente(cas_total[10],self.total_for_customers[11]),calculate_percente(cas_total[11],self.total_for_customers[12]),calculate_percente(cas_total[9]+cas_total[10]+cas_total[11],self.total_for_customers[10]+self.total_for_customers[11]+self.total_for_customers[12]),calculate_percente(sum(cas_total.values()),sum(self.total_for_customers.values()))])

            for c in channels:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    value = self.cas_others_total_for_months[cas]
                    last_year = self.cas_others_total_customer_ly[cas]
                    total_amount = sum(value.values())
                else:
                    value = self.grp_channels["{0}_{1}".format(cas,c)]
                    last_year = self.fn.channel_total_ly(cas,c)
                    total_amount = sum(value)

                data.append([c,calculate_percente(last_year,self.total_customer_ly[cas]),calculate_percente(value[0],cas_total[0]),calculate_percente(value[1],cas_total[1]),calculate_percente(value[2],cas_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_total[0]+cas_total[1]+cas_total[2]),calculate_percente(value[3],cas_total[3]),calculate_percente(value[4],cas_total[4]),calculate_percente(value[5],cas_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_total[3]+cas_total[4]+cas_total[5]),calculate_percente(value[6],cas_total[6]),calculate_percente(value[7],cas_total[7]),calculate_percente(value[8],cas_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_total[6]+cas_total[7]+cas_total[8]),calculate_percente(value[9],cas_total[9]),calculate_percente(value[10],cas_total[10]),calculate_percente(value[11],cas_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_total[9]+cas_total[10]+cas_total[11]),calculate_percente(total_amount,sum(cas_total.values()))])

        return data

    def sov(self):
        sov = {}
        n1 = self.n1_nova_s_total_for_months['N1']
        nova_s = self.n1_nova_s_total_for_months['Nova S']

        # GRAND TOTAL
        grand = [float_with_comma(round(sum(self.total_customer_ly.values())))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            grand += [float_with_comma(round(self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3]))]
        grand += [float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1])), float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2])),float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3]))]

        ct = self.fn.range_by_quarter[self.QUARTER]
        grand += [float_with_comma(round(self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3]))]
        
        grand += [float_with_comma(round(sum(self.total_for_customers.values())))]

        grand += [calculate_percente(sum(self.total_customer_ly.values()),sum(self.total_customer_ly.values()))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            grand += [calculate_percente(self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3],self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3])]
        grand += [calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3])]

        grand += [calculate_percente(self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3],self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3])]

        grand += [calculate_percente(sum(self.total_for_customers.values()),sum(self.total_for_customers.values()))]
        sov['grand'] = grand

        # Nova S
        nova = [float_with_comma(round(self.fn.channel_total_ly_no_customer('Nova S')))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            nova += [float_with_comma(round(nova_s[i]+nova_s[i+1]+nova_s[i+2]))]
        nova += [float_with_comma(round(nova_s[self.fn.range_by_quarter[self.QUARTER]])),float_with_comma(round(nova_s[self.fn.range_by_quarter[self.QUARTER]+1])),float_with_comma(round(nova_s[self.fn.range_by_quarter[self.QUARTER]+2]))]

        nova += [float_with_comma(round(nova_s[ct]+nova_s[ct+1]+nova_s[ct+2]))]
        
        nova += [float_with_comma(round(sum(nova_s.values())))]

        nova += [calculate_percente(self.fn.channel_total_ly_no_customer('Nova S'),sum(self.total_customer_ly.values()))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            nova += [calculate_percente(nova_s[i]+nova_s[i+1]+nova_s[i+2],self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3])]
        nova += [calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]+1],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]+2],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3])]
        
        nova += [calculate_percente(nova_s[ct]+nova_s[ct+1]+nova_s[ct+2],self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3])]

        nova += [calculate_percente(sum(nova_s.values()),sum(self.total_for_customers.values()))]
        sov['nova_s'] = nova

        # N1
        n = [float_with_comma(round(self.fn.channel_total_ly_no_customer('N1')))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            n += [float_with_comma(round(n1[i]+n1[i+1]+n1[i+2]))]
        n += [float_with_comma(round(n1[self.fn.range_by_quarter[self.QUARTER]])),float_with_comma(round(n1[self.fn.range_by_quarter[self.QUARTER]+1])),float_with_comma(round(n1[self.fn.range_by_quarter[self.QUARTER]+2]))]
        
        n += [float_with_comma(round(n1[ct]+n1[ct+1]+n1[ct+2]))]
        n += [float_with_comma(round(sum(n1.values())))]

        n += [calculate_percente(self.fn.channel_total_ly_no_customer('N1'),sum(self.total_customer_ly.values()))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            n += [calculate_percente(n1[i]+n1[i+1]+n1[i+2],self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3])]
        n += [calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]+1],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]+2],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3])]
        
        n += [calculate_percente(n1[ct]+n1[ct+1]+n1[ct+2],self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3])]

        n += [calculate_percente(sum(n1.values()),sum(self.total_for_customers.values()))]
        sov['n1'] = n

        # CAS OTHERS
        co = [float_with_comma(round(sum(self.total_customer_ly.values())-self.fn.channel_total_ly_no_customer('Nova S')-self.fn.channel_total_ly_no_customer('N1')))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            co += [float_with_comma(round(self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3]-nova_s[i]-nova_s[i+1]-nova_s[i+2]-n1[i]-n1[i+1]-n1[i+2]))]
        co += [float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]-nova_s[self.fn.range_by_quarter[self.QUARTER]]-n1[self.fn.range_by_quarter[self.QUARTER]])),float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]-nova_s[self.fn.range_by_quarter[self.QUARTER]+1]-n1[self.fn.range_by_quarter[self.QUARTER]+1])),float_with_comma(round(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3]-nova_s[self.fn.range_by_quarter[self.QUARTER]+2]-n1[self.fn.range_by_quarter[self.QUARTER]+2]))]

        co += [float_with_comma(round(self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3]-nova_s[ct]-nova_s[ct+1]-nova_s[ct+2]-n1[ct]-n1[ct+1]-n1[ct+2]))]
        
        co += [float_with_comma(round(sum(self.total_for_customers.values())-sum(n1.values())-sum(nova_s.values())))]

        co += [calculate_percente(sum(self.total_customer_ly.values())-self.fn.channel_total_ly_no_customer('Nova S')-self.fn.channel_total_ly_no_customer('N1'),sum(self.total_customer_ly.values()))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            co += [calculate_percente(self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3]-nova_s[i]-nova_s[i+1]-nova_s[i+2]-n1[i]-n1[i+1]-n1[i+2],self.total_for_customers[i+1]+self.total_for_customers[i+2]+self.total_for_customers[i+3])]
        co += [calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]-nova_s[self.fn.range_by_quarter[self.QUARTER]]-n1[self.fn.range_by_quarter[self.QUARTER]],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]-nova_s[self.fn.range_by_quarter[self.QUARTER]+1]-n1[self.fn.range_by_quarter[self.QUARTER]+1],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3]-nova_s[self.fn.range_by_quarter[self.QUARTER]+2]-n1[self.fn.range_by_quarter[self.QUARTER]+2],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3])]

        co += [calculate_percente(self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3]-nova_s[ct]-nova_s[ct+1]-nova_s[ct+2]-n1[ct]-n1[ct+1]-n1[ct+2],self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3])]
        
        co += [calculate_percente(sum(self.total_for_customers.values())-sum(n1.values())-sum(nova_s.values()),sum(self.total_for_customers.values()))]
        sov['cas_others'] = co

        for cas in self.fn.customers:
            cas_total = self.grp_customers_v[cas]
            s = [float_with_comma(round(self.total_customer_ly[cas]))]
            for q in range(1,self.QUARTER):
                i2 = self.fn.range_by_quarter[q]
                s += [float_with_comma(round(cas_total[i2]+cas_total[i2+1]+cas_total[i2+2]))]
            s += [float_with_comma(round(cas_total[self.fn.range_by_quarter[self.QUARTER]])),float_with_comma(round(cas_total[self.fn.range_by_quarter[self.QUARTER]+1])),float_with_comma(round(cas_total[self.fn.range_by_quarter[self.QUARTER]+2])),float_with_comma(round(cas_total[ct]+cas_total[ct+1]+cas_total[ct+2])), float_with_comma(round(sum(cas_total.values())))]

            s += [calculate_percente(self.total_customer_ly[cas],sum(self.total_customer_ly.values()))]
            for q in range(1,self.QUARTER):
                i2 = self.fn.range_by_quarter[q]
                s += [calculate_percente(cas_total[i2]+cas_total[i2+1]+cas_total[i2+2],self.total_for_customers[i2+1]+self.total_for_customers[i2+2]+self.total_for_customers[i2+3])]
            s += [calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]+1],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(cas_total[self.fn.range_by_quarter[self.QUARTER]+2],self.total_for_customers[self.fn.range_by_quarter[self.QUARTER]+3]),calculate_percente(cas_total[ct]+cas_total[ct+1]+cas_total[ct+2],self.total_for_customers[ct+1]+self.total_for_customers[ct+2]+self.total_for_customers[ct+3]),calculate_percente(sum(cas_total.values()),sum(self.total_for_customers.values()))]
            
            sov["{0}_total".format(cas)] = s

            for c in ['N1', 'Nova S', 'CAS Others (w/o SK, Nova Sport, Brainz)']:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    last_year = self.cas_others_total_customer_ly[cas]
                    value = self.cas_others_total_for_months[cas]
                    total_amount = sum(value.values())
                else:
                    value = self.grp_channels["{0}_{1}".format(cas,c)]
                    last_year = self.fn.channel_total_ly(cas,c)
                    total_amount = sum(value)

                ch = [float_with_comma(round(last_year))]
                for q in range(1,self.QUARTER):
                    ch += [float_with_comma(round(value[ct]+value[ct+1]+value[ct+2]))]
                ch += [float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]])),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+1])),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+2])),float_with_comma(round(value[ct]+value[ct+1]+value[ct+2])),float_with_comma(round(total_amount))]

                ch += [calculate_percente(last_year,self.total_customer_ly[cas])]
                for q in range(1,self.QUARTER):
                    ch += [calculate_percente(value[ct]+value[ct+1]+value[ct+2],cas_total[ct]+cas_total[ct+1]+cas_total[ct+2])]
                ch += [calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]],cas_total[self.fn.range_by_quarter[self.QUARTER]]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+1],cas_total[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+2],cas_total[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(value[ct]+value[ct+1]+value[ct+2],cas_total[ct]+cas_total[ct+1]+cas_total[ct+2]),calculate_percente(total_amount,sum(cas_total.values()))]
                    
                sov["{0}_{1}".format(cas,c)] = ch

        return sov

    def table_18_50(self):
        res = self.total_table()
        res[0] = ['18-50']
        return res

    def sov_18_50(self):
        res = self.sov_table()
        res[0] = ['SOV 18-50']
        return res

    def shr_sov_total_table(self):
        data = [['Total']]
        data.append(["SHR% of Audience TOTAL 4+ (08-24 h)"])
        data.append(['Row Labels'] + self.headers)

        values = self.fn.non_shr_channel_value(13,31)

        a1 = ['CAS Media']
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a1 += ['{0}%'.format(round(total,2))]

        a2 = ['CAS Media AGB']
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a2 += ['{0}%'.format(round(total,2))]

        data.append(a2)
        data.append(a1)

        for d in self.fn.non_shr_channel_value_2(13,31):
            if d[0] in self.fn.dm_pool_shr_channels:
                data.append(d)
        return data

    def shr_18_50_total_table(self):
        data = [['18-50']]
        data.append(["SHR% of Audience 18-50 (08-24 h)"])
        data.append(['Row Labels'] + self.headers)

        values = self.fn.non_shr_channel_value(45,63)

        a1 = ['CAS Media']
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a1 += ['{0}%'.format(round(total,2))]

        a2 = ['CAS Media AGB']
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a2 += ['{0}%'.format(round(total,2))]

        data.append(a2)
        data.append(a1)

        for d in self.fn.non_shr_channel_value_2(45,63):
            if d[0] in self.fn.dm_pool_shr_channels:
                data.append(d)
        return data


    def kontrolni_duration(self):
        data = [['PRIPREMNI ZA DURATION']]
        data.append(['DURATION'])
        data.append(['Row Labels']+self.headers[1:])

        n1 = self.n1_nova_s_duration['N1']
        nova_s = self.n1_nova_s_duration['Nova S']

        data.append(['Cas Media',calculate_days_float(self.all_duration[1]),calculate_days_float(self.all_duration[2]),calculate_days_float(self.all_duration[3]),calculate_days_float(self.all_duration[1]+self.all_duration[2]+self.all_duration[3]),calculate_days_float(self.all_duration[4]),calculate_days_float(self.all_duration[5]),calculate_days_float(self.all_duration[6]),calculate_days_float(self.all_duration[4]+self.all_duration[5]+self.all_duration[6]),calculate_days_float(self.all_duration[7]),calculate_days_float(self.all_duration[8]),calculate_days_float(self.all_duration[9]),calculate_days_float(self.all_duration[7]+self.all_duration[8]+self.all_duration[9]),calculate_days_float(self.all_duration[10]),calculate_days_float(self.all_duration[11]),calculate_days_float(self.all_duration[12]),calculate_days_float(self.all_duration[10]+self.all_duration[11]+self.all_duration[12]),calculate_days_float(self.fn.sum_delta_time(self.all_duration.values()))])

        data.append(['Nova S',calculate_days_float(nova_s[0]),calculate_days_float(nova_s[1]),calculate_days_float(nova_s[2]),calculate_days_float(nova_s[0]+nova_s[1]+nova_s[2]),calculate_days_float(nova_s[3]),calculate_days_float(nova_s[4]),calculate_days_float(nova_s[5]),calculate_days_float(nova_s[3]+nova_s[4]+nova_s[5]),calculate_days_float(nova_s[6]),calculate_days_float(nova_s[7]),calculate_days_float(nova_s[8]),calculate_days_float(nova_s[6]+nova_s[7]+nova_s[8]),calculate_days_float(nova_s[9]),calculate_days_float(nova_s[10]),calculate_days_float(nova_s[11]),calculate_days_float(nova_s[9]+nova_s[10]+nova_s[11]),calculate_days_float(self.fn.sum_delta_time(nova_s.values()))])

        data.append(['N1',calculate_days_float(n1[0]),calculate_days_float(n1[1]),calculate_days_float(n1[2]),calculate_days_float(n1[0]+n1[1]+n1[2]),calculate_days_float(n1[3]),calculate_days_float(n1[4]),calculate_days_float(n1[5]),calculate_days_float(n1[3]+n1[4]+n1[5]),calculate_days_float(n1[6]),calculate_days_float(n1[7]),calculate_days_float(n1[8]),calculate_days_float(n1[6]+n1[7]+n1[8]),calculate_days_float(n1[9]),calculate_days_float(n1[10]),calculate_days_float(n1[11]),calculate_days_float(n1[9]+n1[10]+n1[11]),calculate_days_float(self.fn.sum_delta_time(n1.values()))])

        data.append(['CAS Others (w/o SK, Nova Sport, Brainz)',calculate_days_float(self.all_duration[1]-nova_s[0]-n1[0]),calculate_days_float(self.all_duration[2]-nova_s[1]-n1[1]),calculate_days_float(self.all_duration[3]-nova_s[2]-n1[2]),
        calculate_days_float(self.all_duration[1]+self.all_duration[2]+self.all_duration[3]-nova_s[0]-nova_s[1]-nova_s[2]-n1[0]-n1[1]-n1[2]),calculate_days_float(self.all_duration[4]-nova_s[3]-n1[3]),
        calculate_days_float(self.all_duration[5]-nova_s[4]-n1[4]),calculate_days_float(self.all_duration[6]-nova_s[5]-n1[5]),calculate_days_float(self.all_duration[4]+self.all_duration[5]+self.all_duration[6]-nova_s[3]-nova_s[4]-nova_s[5]-n1[3]-n1[4]-n1[5]),calculate_days_float(self.all_duration[7]-nova_s[6]-n1[6]),calculate_days_float(self.all_duration[8]-nova_s[7]-n1[7]),calculate_days_float(self.all_duration[9]-nova_s[8]-n1[8]),calculate_days_float(self.all_duration[7]+self.all_duration[8]+self.all_duration[9]-nova_s[6]-nova_s[7]-nova_s[8]-n1[6]-n1[7]-n1[8]),calculate_days_float(self.all_duration[10]-nova_s[9]-n1[9]),calculate_days_float(self.all_duration[11]-nova_s[10]-n1[10]),calculate_days_float(self.all_duration[12]-nova_s[11]-n1[11]),calculate_days_float(self.all_duration[10]+self.all_duration[11]+self.all_duration[12]-nova_s[9]-nova_s[10]-nova_s[11]-n1[9]-n1[10]-n1[11]),calculate_days_float(self.fn.sum_delta_time(self.all_duration.values())-self.fn.sum_delta_time(n1.values())-self.fn.sum_delta_time(nova_s.values()))])

        channels = self.fn.cas_other_channels
        for cas in self.fn.customers:
            cas_total = self.cas_duration[cas]

            data.append([cas,calculate_days_float(cas_total[0]),calculate_days_float(cas_total[1]),calculate_days_float(cas_total[2]),calculate_days_float(cas_total[0]+cas_total[1]+cas_total[2]),calculate_days_float(cas_total[3]),calculate_days_float(cas_total[4]),calculate_days_float(cas_total[5]),calculate_days_float(cas_total[3]+cas_total[4]+cas_total[5]),calculate_days_float(cas_total[6]),calculate_days_float(cas_total[7]),calculate_days_float(cas_total[8]),calculate_days_float(cas_total[6]+cas_total[7]+cas_total[8]),calculate_days_float(cas_total[9]),calculate_days_float(cas_total[10]),calculate_days_float(cas_total[11]),calculate_days_float(cas_total[9]+cas_total[10]+cas_total[11]),calculate_days_float(self.fn.sum_delta_time(cas_total.values()))])

            for c in channels:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    value = self.cas_others_duration[cas]
                    total_amount = self.fn.sum_delta_time(value.values())
                else:
                    value = self.c_duration["{0}_{1}".format(cas,c)]
                    total_amount = self.fn.sum_delta_time(value)

                data.append([c,calculate_days_float(value[0]),calculate_days_float(value[1]),calculate_days_float(value[2]),calculate_days_float(value[0]+value[1]+value[2]),calculate_days_float(value[3]),calculate_days_float(value[4]),calculate_days_float(value[5]),calculate_days_float(value[3]+value[4]+value[5]),calculate_days_float(value[6]),calculate_days_float(value[7]),calculate_days_float(value[8]),calculate_days_float(value[6]+value[7]+value[8]),calculate_days_float(value[9]),calculate_days_float(value[10]),calculate_days_float(value[11]),calculate_days_float(value[9]+value[10]+value[11]),calculate_days_float(total_amount)])
        
        data.append(['TOTAL',calculate_days_float(self.duration_total_by_month[0]),calculate_days_float(self.duration_total_by_month[1]),calculate_days_float(self.duration_total_by_month[2]),calculate_days_float(self.duration_total_by_month[0]+self.duration_total_by_month[1]+self.duration_total_by_month[2]),calculate_days_float(self.duration_total_by_month[3]),calculate_days_float(self.duration_total_by_month[4]),calculate_days_float(self.duration_total_by_month[5]),calculate_days_float(self.duration_total_by_month[3]+self.duration_total_by_month[4]+self.duration_total_by_month[5]),calculate_days_float(self.duration_total_by_month[6]),calculate_days_float(self.duration_total_by_month[7]),calculate_days_float(self.duration_total_by_month[8]),calculate_days_float(self.duration_total_by_month[6]+self.duration_total_by_month[7]+self.duration_total_by_month[8]),calculate_days_float(self.duration_total_by_month[9]),calculate_days_float(self.duration_total_by_month[10]),calculate_days_float(self.duration_total_by_month[11]),calculate_days_float(self.duration_total_by_month[9]+self.duration_total_by_month[10]+self.duration_total_by_month[11]),calculate_days_float(self.fn.sum_delta_time(self.duration_total_by_month))])

        data.append(['TOTAL Provera',calculate_days_float(self.all_duration[1]),calculate_days_float(self.all_duration[2]),calculate_days_float(self.all_duration[3]),calculate_days_float(self.all_duration[1]+self.all_duration[2]+self.all_duration[3]),calculate_days_float(self.all_duration[4]),calculate_days_float(self.all_duration[5]),calculate_days_float(self.all_duration[6]),calculate_days_float(self.all_duration[4]+self.all_duration[5]+self.all_duration[6]),calculate_days_float(self.all_duration[7]),calculate_days_float(self.all_duration[8]),calculate_days_float(self.all_duration[9]),calculate_days_float(self.all_duration[7]+self.all_duration[8]+self.all_duration[9]),calculate_days_float(self.all_duration[10]),calculate_days_float(self.all_duration[11]),calculate_days_float(self.all_duration[12]),calculate_days_float(self.all_duration[10]+self.all_duration[11]+self.all_duration[12]),calculate_days_float(self.fn.sum_delta_time(self.all_duration.values()))])

        return data


    def duration(self,cover_duration):
        data = [['DURATION']]
        data.append(['Row Labels']+self.headers)

        n1 = self.n1_nova_s_duration['N1']
        nova_s = self.n1_nova_s_duration['Nova S']

        data.append(['CAS Media','{0} %'.format(round(self.duration_ly['CAS Media'],2)),calculate_percente(self.all_duration[1],self.duration_total_by_month[0]),calculate_percente(self.all_duration[2],self.duration_total_by_month[1]),calculate_percente(self.all_duration[3],self.duration_total_by_month[2]),calculate_percente(self.all_duration[1]+self.all_duration[2]+self.all_duration[3],self.duration_total_by_month[0]+self.duration_total_by_month[1]+self.duration_total_by_month[2]),calculate_percente(self.all_duration[4],self.duration_total_by_month[3]),calculate_percente(self.all_duration[5],self.duration_total_by_month[4]),calculate_percente(self.all_duration[6],self.duration_total_by_month[5]),calculate_percente(self.all_duration[4]+self.all_duration[5]+self.all_duration[6],self.duration_total_by_month[3]+self.duration_total_by_month[4]+self.duration_total_by_month[5]),calculate_percente(self.all_duration[7],self.duration_total_by_month[6]),calculate_percente(self.all_duration[8],self.duration_total_by_month[7]),calculate_percente(self.all_duration[9],self.duration_total_by_month[8]),calculate_percente(self.all_duration[7]+self.all_duration[8]+self.all_duration[9],self.duration_total_by_month[6]+self.duration_total_by_month[7]+self.duration_total_by_month[8]),calculate_percente(self.all_duration[10],self.duration_total_by_month[9]),calculate_percente(self.all_duration[11],self.duration_total_by_month[10]),calculate_percente(self.all_duration[12],self.duration_total_by_month[11]),calculate_percente(self.all_duration[10]+self.all_duration[11]+self.all_duration[12],self.duration_total_by_month[9]+self.duration_total_by_month[10]+self.duration_total_by_month[11]),calculate_percente(self.fn.sum_delta_time(self.all_duration.values()),self.fn.sum_delta_time(self.duration_total_by_month))])


        data.append(['Nova S','{0} %'.format(round(self.duration_ly['Nova S'],2)),calculate_percente(nova_s[0],self.all_duration[1]),calculate_percente(nova_s[1],self.all_duration[2]),calculate_percente(nova_s[2],self.all_duration[3]),calculate_percente(nova_s[0]+nova_s[1]+nova_s[2],self.all_duration[1]+self.all_duration[2]+self.all_duration[3]),calculate_percente(nova_s[3],self.all_duration[4]),calculate_percente(nova_s[4],self.all_duration[5]),calculate_percente(nova_s[5],self.all_duration[6]),calculate_percente(nova_s[3]+nova_s[4]+nova_s[5],self.all_duration[4]+self.all_duration[5]+self.all_duration[6]),calculate_percente(nova_s[6],self.all_duration[7]),calculate_percente(nova_s[7],self.all_duration[8]),calculate_percente(nova_s[8],self.all_duration[9]),calculate_percente(nova_s[6]+nova_s[7]+nova_s[8],self.all_duration[7]+self.all_duration[8]+self.all_duration[9]),calculate_percente(nova_s[9],self.all_duration[10]),calculate_percente(nova_s[10],self.all_duration[11]),calculate_percente(nova_s[11],self.all_duration[12]),calculate_percente(nova_s[9]+nova_s[10]+nova_s[11],self.all_duration[10]+self.all_duration[11]+self.all_duration[12]),calculate_percente(self.fn.sum_delta_time(nova_s.values()),self.fn.sum_delta_time(self.all_duration.values()))])

        data.append(['N1','{0} %'.format(round(self.duration_ly['N1'],2)),calculate_percente(n1[0],self.all_duration[1]),calculate_percente(n1[1],self.all_duration[2]),calculate_percente(n1[2],self.all_duration[3]),calculate_percente(n1[0]+n1[1]+n1[2],self.all_duration[1]+self.all_duration[2]+self.all_duration[3]),calculate_percente(n1[3],self.all_duration[4]),calculate_percente(n1[4],self.all_duration[5]),calculate_percente(n1[5],self.all_duration[6]),calculate_percente(n1[3]+n1[4]+n1[5],self.all_duration[4]+self.all_duration[5]+self.all_duration[6]),calculate_percente(n1[6],self.all_duration[7]),calculate_percente(n1[7],self.all_duration[8]),calculate_percente(n1[8],self.all_duration[9]),calculate_percente(n1[6]+n1[7]+n1[8],self.all_duration[7]+self.all_duration[8]+self.all_duration[9]),calculate_percente(n1[9],self.all_duration[10]),calculate_percente(n1[10],self.all_duration[11]),calculate_percente(n1[11],self.all_duration[12]),calculate_percente(n1[9]+n1[10]+n1[11],self.all_duration[10]+self.all_duration[11]+self.all_duration[12]),calculate_percente(self.fn.sum_delta_time(n1.values()),self.fn.sum_delta_time(self.all_duration.values()))])


        data.append(['CAS Others (w/o SK, Nova Sport, Brainz)','{0} %'.format(round(self.duration_ly['CAS Others (w/o SK, Nova Sport, Brainz)'],2)),calculate_percente(self.all_duration[1]-nova_s[0]-n1[0],self.all_duration[1]),calculate_percente(self.all_duration[2]-nova_s[1]-n1[1],self.all_duration[2]),calculate_percente(self.all_duration[3]-nova_s[2]-n1[2],self.all_duration[3]),calculate_percente(self.all_duration[1]+self.all_duration[2]+self.all_duration[3]-nova_s[0]-nova_s[1]-nova_s[2]-n1[0]-n1[1]-n1[2],self.all_duration[1]+self.all_duration[2]+self.all_duration[3]),calculate_percente(self.all_duration[4]-nova_s[3]-n1[3],self.all_duration[4]),calculate_percente(self.all_duration[5]-nova_s[4]-n1[4],self.all_duration[5]),calculate_percente(self.all_duration[6]-nova_s[5]-n1[5],self.all_duration[6]),calculate_percente(self.all_duration[4]+self.all_duration[5]+self.all_duration[6]-nova_s[3]-nova_s[4]-nova_s[5]-n1[3]-n1[4]-n1[5],self.all_duration[4]+self.all_duration[5]+self.all_duration[6]),calculate_percente(self.all_duration[7]-nova_s[6]-n1[6],self.all_duration[7]),calculate_percente(self.all_duration[8]-nova_s[7]-n1[7],self.all_duration[8]),calculate_percente(self.all_duration[9]-nova_s[8]-n1[8],self.all_duration[9]),calculate_percente(self.all_duration[7]+self.all_duration[8]+self.all_duration[9]-nova_s[6]-nova_s[7]-nova_s[8]-n1[6]-n1[7]-n1[8],self.all_duration[7]+self.all_duration[8]+self.all_duration[9]),calculate_percente(self.all_duration[10]-nova_s[9]-n1[9],self.all_duration[10]),calculate_percente(self.all_duration[11]-nova_s[10]-n1[10],self.all_duration[11]),calculate_percente(self.all_duration[12]-nova_s[11]-n1[11],self.all_duration[12]),calculate_percente(n1[9]+n1[10]+n1[11],self.all_duration[10]+self.all_duration[11]+self.all_duration[12]),calculate_percente(self.fn.sum_delta_time(self.all_duration.values())-self.fn.sum_delta_time(n1.values())-self.fn.sum_delta_time(nova_s.values()),self.fn.sum_delta_time(self.all_duration.values()))])

        channels = self.fn.cas_other_channels
        for cas in self.fn.customers:
            cas_total = cover_duration[cas]
            cas_customer_total = self.cas_duration[cas]

            data.append([cas,cas_total[0],cas_total[1],cas_total[2],cas_total[3],cas_total[4],cas_total[5],cas_total[6],cas_total[7],cas_total[8],cas_total[9],cas_total[10],cas_total[11],cas_total[12],cas_total[13],cas_total[14],cas_total[15],cas_total[16],cas_total[17]])

            for c in channels:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    value = self.cas_others_duration[cas]
                    total_amount = self.fn.sum_delta_time(value.values())
                else:
                    value = self.c_duration["{0}_{1}".format(cas,c)]
                    total_amount = self.fn.sum_delta_time(value)

                data.append([c,'{0} %'.format(round(self.duration_ly["{0}_{1}".format(cas,c)],2)),calculate_percente(value[0],cas_customer_total[0]),calculate_percente(value[1],cas_customer_total[1]),calculate_percente(value[2],cas_customer_total[2]),calculate_percente(value[0]+value[1]+value[2],cas_customer_total[0]+cas_customer_total[1]+cas_customer_total[2]),calculate_percente(value[3],cas_customer_total[3]),calculate_percente(value[4],cas_customer_total[4]),calculate_percente(value[5],cas_customer_total[5]),calculate_percente(value[3]+value[4]+value[5],cas_customer_total[3]+cas_customer_total[4]+cas_customer_total[5]),calculate_percente(value[6],cas_customer_total[6]),calculate_percente(value[7],cas_customer_total[7]),calculate_percente(value[8],cas_customer_total[8]),calculate_percente(value[6]+value[7]+value[8],cas_customer_total[6]+cas_customer_total[7]+cas_customer_total[8]),calculate_percente(value[9],cas_customer_total[9]),calculate_percente(value[10],cas_customer_total[10]),calculate_percente(value[11],cas_customer_total[11]),calculate_percente(value[9]+value[10]+value[11],cas_customer_total[9]+cas_customer_total[10]+cas_customer_total[11]),calculate_percente(total_amount,self.fn.sum_delta_time(cas_customer_total.values()))])

        return data


    def soi_total_table(self,soi_header):
        new_header = []
        for h in self.headers:
            new_header.append(h)
            new_header.append('')

        data = [['COST, Net Eur; SOI%']]
        data.append(['Row Labels']+new_header)

        data.append(['DM Pool',float_with_comma(round(soi_header[0])),'',float_with_comma(round(soi_header[1])),'',float_with_comma(round(soi_header[2])),'',float_with_comma(round(soi_header[3])),'',float_with_comma(round(soi_header[4])),'',float_with_comma(round(soi_header[5])),'',float_with_comma(round(soi_header[6])),'',float_with_comma(round(soi_header[7])),'',float_with_comma(round(soi_header[8])),'',float_with_comma(round(soi_header[9])),'',float_with_comma(round(soi_header[10])),'',float_with_comma(round(soi_header[11])),'',float_with_comma(round(soi_header[12])),'',float_with_comma(round(soi_header[13])),'',float_with_comma(round(soi_header[14])),'',float_with_comma(round(soi_header[15])),'',float_with_comma(round(soi_header[16])),'',float_with_comma(round(soi_header[17]))])

        data.append(['CAS Media',float_with_comma(round(sum(self.soi_channel_total_ly.values()))),calculate_percente_round_1(sum(self.soi_channel_total_ly.values()),soi_header[0]),float_with_comma(round(self.cas_media_total[0])),calculate_percente_round_1(self.cas_media_total[0],soi_header[1]),float_with_comma(round(self.cas_media_total[1])),calculate_percente_round_1(self.cas_media_total[1],soi_header[2]),float_with_comma(round(self.cas_media_total[2])),calculate_percente_round_1(self.cas_media_total[2],soi_header[3]),float_with_comma(round(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2])),calculate_percente_round_1(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2],soi_header[4]),float_with_comma(round(self.cas_media_total[3])),calculate_percente_round_1(self.cas_media_total[3],soi_header[5]),float_with_comma(round(self.cas_media_total[4])),calculate_percente_round_1(self.cas_media_total[4],soi_header[6]),float_with_comma(round(self.cas_media_total[5])),calculate_percente_round_1(self.cas_media_total[5],soi_header[7]),float_with_comma(round(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5])),calculate_percente_round_1(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5],soi_header[8]),float_with_comma(round(self.cas_media_total[6])),calculate_percente_round_1(self.cas_media_total[6],soi_header[9]),float_with_comma(round(self.cas_media_total[7])),calculate_percente_round_1(self.cas_media_total[7],soi_header[10]),float_with_comma(round(self.cas_media_total[8])),calculate_percente_round_1(self.cas_media_total[8],soi_header[11]),float_with_comma(round(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8])),calculate_percente_round_1(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8],soi_header[12]),float_with_comma(round(self.cas_media_total[9])),calculate_percente_round_1(self.cas_media_total[9],soi_header[13]),float_with_comma(round(self.cas_media_total[10])),calculate_percente_round_1(self.cas_media_total[10],soi_header[14]),float_with_comma(round(self.cas_media_total[11])),calculate_percente_round_1(self.cas_media_total[11],soi_header[15]),float_with_comma(round(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11])),calculate_percente_round_1(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11],soi_header[16]),float_with_comma(round(sum(self.cas_media_total.values()))),calculate_percente_round_1(sum(self.cas_media_total.values()),soi_header[17])])


        data.append(['CAS Media AGB',float_with_comma(round(self.cas_media_agb_ly)),calculate_percente_round_1(self.cas_media_agb_ly,sum(self.soi_channel_total_ly.values())),float_with_comma(round(self.cas_media_agb[0])),calculate_percente_round_1(self.cas_media_agb[0],self.cas_media_total[0]),float_with_comma(round(self.cas_media_agb[1])),calculate_percente_round_1(self.cas_media_agb[1],self.cas_media_total[1]),float_with_comma(round(self.cas_media_agb[2])),calculate_percente_round_1(self.cas_media_agb[2],self.cas_media_total[2]),float_with_comma(round(self.cas_media_agb[0]+self.cas_media_agb[1]+self.cas_media_agb[2])),calculate_percente_round_1(self.cas_media_agb[0]+self.cas_media_agb[1]+self.cas_media_agb[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),float_with_comma(round(self.cas_media_agb[3])),calculate_percente_round_1(self.cas_media_agb[3],self.cas_media_total[3]),float_with_comma(round(self.cas_media_agb[4])),calculate_percente_round_1(self.cas_media_agb[4],self.cas_media_total[4]),float_with_comma(round(self.cas_media_agb[5])),calculate_percente_round_1(self.cas_media_agb[5],self.cas_media_total[5]),float_with_comma(round(self.cas_media_agb[3]+self.cas_media_agb[4]+self.cas_media_agb[5])),calculate_percente_round_1(self.cas_media_agb[3]+self.cas_media_agb[4]+self.cas_media_agb[5],self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5]),float_with_comma(round(self.cas_media_agb[6])),calculate_percente_round_1(self.cas_media_agb[6],self.cas_media_total[6]),float_with_comma(round(self.cas_media_agb[7])),calculate_percente_round_1(self.cas_media_agb[7],self.cas_media_total[7]),float_with_comma(round(self.cas_media_agb[8])),calculate_percente_round_1(self.cas_media_agb[8],self.cas_media_total[8]),float_with_comma(round(self.cas_media_agb[6]+self.cas_media_agb[7]+self.cas_media_agb[8])),calculate_percente_round_1(self.cas_media_agb[6]+self.cas_media_agb[7]+self.cas_media_agb[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),float_with_comma(round(self.cas_media_agb[9])),calculate_percente_round_1(self.cas_media_agb[9],self.cas_media_total[9]),float_with_comma(round(self.cas_media_agb[10])),calculate_percente_round_1(self.cas_media_agb[10],self.cas_media_total[10]),float_with_comma(round(self.cas_media_agb[11])),calculate_percente_round_1(self.cas_media_agb[11],self.cas_media_total[11]),float_with_comma(round(self.cas_media_agb[9]+self.cas_media_agb[10]+self.cas_media_agb[11])),calculate_percente_round_1(self.cas_media_agb[9]+self.cas_media_agb[10]+self.cas_media_agb[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),float_with_comma(round(sum(self.cas_media_agb.values()))),calculate_percente_round_1(sum(self.cas_media_agb.values()),sum(self.cas_media_total.values()))])

        channels = self.fn.soi_e2_channels()

        for c in channels:
              value = self.soi_total2[c]

              data.append([c,float_with_comma(round(self.soi_channel_total_ly[c])),calculate_percente_round_1(self.soi_channel_total_ly[c],sum(self.soi_channel_total_ly.values())),float_with_comma(round(value[0])),calculate_percente_round_1(value[0],self.cas_media_total[0]),float_with_comma(round(value[1])),calculate_percente_round_1(value[1],self.cas_media_total[1]),float_with_comma(round(value[2])),calculate_percente_round_1(value[2],self.cas_media_total[2]),float_with_comma(round(value[0]+value[1]+value[2])),calculate_percente_round_1(value[0]+value[1]+value[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),float_with_comma(round(value[3])),calculate_percente_round_1(value[3],self.cas_media_total[3]),float_with_comma(round(value[4])),calculate_percente_round_1(value[4],self.cas_media_total[4]),float_with_comma(round(value[5])),calculate_percente_round_1(value[5],self.cas_media_total[5]),float_with_comma(round(value[3]+value[4]+value[5])),calculate_percente_round_1(value[3]+value[4]+value[5],self.cas_media_total[4]+self.cas_media_total[5]+self.cas_media_total[6]),float_with_comma(round(value[6])),calculate_percente_round_1(value[6],self.cas_media_total[6]),float_with_comma(round(value[7])),calculate_percente_round_1(value[7],self.cas_media_total[7]),float_with_comma(round(value[8])),calculate_percente_round_1(value[8],self.cas_media_total[8]),float_with_comma(round(value[6]+value[7]+value[8])),calculate_percente_round_1(value[6]+value[7]+value[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),float_with_comma(round(value[9])),calculate_percente_round_1(value[9],self.cas_media_total[9]),float_with_comma(round(value[10])),calculate_percente_round_1(value[10],self.cas_media_total[10]),float_with_comma(round(value[11])),calculate_percente_round_1(value[11],self.cas_media_total[11]),float_with_comma(round(value[9]+value[10]+value[11])),calculate_percente_round_1(value[9]+value[10]+value[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),float_with_comma(round(sum(value))),calculate_percente_round_1(sum(value),sum(self.cas_media_total.values()))])

        co = ['CAS Others (w/o SK, Nova Sport, Brainz)',float_with_comma(round(self.soi_cas_others_channels_ly)),calculate_percente_round_1(self.soi_cas_others_channels_ly,sum(self.soi_channel_total_ly.values())),float_with_comma(round(self.soi_cas_others[0])),calculate_percente_round_1(self.soi_cas_others[0],self.cas_media_total[0]),float_with_comma(round(self.soi_cas_others[1])),calculate_percente_round_1(self.soi_cas_others[1],self.cas_media_total[1]),float_with_comma(round(self.soi_cas_others[2])),calculate_percente_round_1(self.soi_cas_others[2],self.cas_media_total[2]),float_with_comma(round(self.soi_cas_others[0]+self.soi_cas_others[1]+self.soi_cas_others[2])),calculate_percente_round_1(self.soi_cas_others[0]+self.soi_cas_others[1]+self.soi_cas_others[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),float_with_comma(round(self.soi_cas_others[3])),calculate_percente_round_1(self.soi_cas_others[3],self.cas_media_total[3]),float_with_comma(round(self.soi_cas_others[4])),calculate_percente_round_1(self.soi_cas_others[4],self.cas_media_total[4]),float_with_comma(round(self.soi_cas_others[5])),calculate_percente_round_1(self.soi_cas_others[5],self.cas_media_total[5]),float_with_comma(round(self.soi_cas_others[3]+self.soi_cas_others[4]+self.soi_cas_others[5])),calculate_percente_round_1(self.soi_cas_others[3]+self.soi_cas_others[4]+self.soi_cas_others[5],self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5]),float_with_comma(round(self.soi_cas_others[6])),calculate_percente_round_1(self.soi_cas_others[6],self.cas_media_total[6]),float_with_comma(round(self.soi_cas_others[7])),calculate_percente_round_1(self.soi_cas_others[7],self.cas_media_total[7]),float_with_comma(round(self.soi_cas_others[8])),calculate_percente_round_1(self.soi_cas_others[8],self.cas_media_total[8]),float_with_comma(round(self.soi_cas_others[6]+self.soi_cas_others[7]+self.soi_cas_others[8])),calculate_percente_round_1(self.soi_cas_others[6]+self.soi_cas_others[7]+self.soi_cas_others[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),float_with_comma(round(self.soi_cas_others[9])),calculate_percente_round_1(self.soi_cas_others[9],self.cas_media_total[9]),float_with_comma(round(self.soi_cas_others[10])),calculate_percente_round_1(self.soi_cas_others[10],self.cas_media_total[10]),float_with_comma(round(self.soi_cas_others[11])),calculate_percente_round_1(self.soi_cas_others[11],self.cas_media_total[11]),float_with_comma(round(self.soi_cas_others[9]+self.soi_cas_others[10]+self.soi_cas_others[11])),calculate_percente_round_1(self.soi_cas_others[9]+self.soi_cas_others[10]+self.soi_cas_others[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),float_with_comma(round(sum(self.soi_cas_others.values()))),calculate_percente_round_1(sum(self.soi_cas_others.values()),sum(self.cas_media_total.values()))]


        co2 = ['CAS "non AGB" (SK, Nova Sport, Brainz)',float_with_comma(round(self.cas_media_channels_ly)),calculate_percente_round_1(self.cas_media_channels_ly,sum(self.soi_channel_total_ly.values())),float_with_comma(round(self.soi_agb_channal[0])),calculate_percente_round_1(self.soi_agb_channal[0],self.cas_media_total[0]),float_with_comma(round(self.soi_agb_channal[1])),calculate_percente_round_1(self.soi_agb_channal[1],self.cas_media_total[1]),float_with_comma(round(self.soi_agb_channal[2])),calculate_percente_round_1(self.soi_agb_channal[2],self.cas_media_total[2]),float_with_comma(round(self.soi_agb_channal[0]+self.soi_agb_channal[1]+self.soi_agb_channal[2])),calculate_percente_round_1(self.soi_agb_channal[0]+self.soi_agb_channal[1]+self.soi_agb_channal[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),float_with_comma(round(self.soi_agb_channal[3])),calculate_percente_round_1(self.soi_agb_channal[3],self.cas_media_total[3]),float_with_comma(round(self.soi_agb_channal[4])),calculate_percente_round_1(self.soi_agb_channal[4],self.cas_media_total[4]),float_with_comma(round(self.soi_agb_channal[5])),calculate_percente_round_1(self.soi_agb_channal[5],self.cas_media_total[5]),float_with_comma(round(self.soi_agb_channal[3]+self.soi_agb_channal[4]+self.soi_agb_channal[5])),calculate_percente_round_1(self.soi_agb_channal[3]+self.soi_agb_channal[4]+self.soi_agb_channal[5],self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5]),float_with_comma(round(self.soi_agb_channal[6])),calculate_percente_round_1(self.soi_agb_channal[6],self.cas_media_total[6]),float_with_comma(round(self.soi_agb_channal[7])),calculate_percente_round_1(self.soi_agb_channal[7],self.cas_media_total[7]),float_with_comma(round(self.soi_agb_channal[8])),calculate_percente_round_1(self.soi_agb_channal[8],self.cas_media_total[8]),float_with_comma(round(self.soi_agb_channal[6]+self.soi_agb_channal[7]+self.soi_agb_channal[8])),calculate_percente_round_1(self.soi_agb_channal[6]+self.soi_agb_channal[7]+self.soi_agb_channal[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),float_with_comma(round(self.soi_agb_channal[9])),calculate_percente_round_1(self.soi_agb_channal[9],self.cas_media_total[9]),float_with_comma(round(self.soi_agb_channal[10])),calculate_percente_round_1(self.soi_agb_channal[10],self.cas_media_total[10]),float_with_comma(round(self.soi_agb_channal[11])),calculate_percente_round_1(self.soi_agb_channal[11],self.cas_media_total[11]),float_with_comma(round(self.soi_agb_channal[9]+self.soi_agb_channal[10]+self.soi_agb_channal[11])),calculate_percente_round_1(self.soi_agb_channal[9]+self.soi_agb_channal[10]+self.soi_agb_channal[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),float_with_comma(round(sum(self.soi_agb_channal.values()))),calculate_percente_round_1(sum(self.soi_agb_channal.values()),sum(self.cas_media_total.values()))]

        data.insert(7,co)
        data.insert(-3,co2)

        return data


    def cpp_4_plus_table(self):
        data = [['Avg CPP 4+']]
        data.append(['Row Labels']+self.headers)

        cas_total = self.grp_customers_v['DM Pool']

        data.append(['CAS Media',div(self.cas_media_agb_ly,self.total_customer_ly['DM Pool']),div(self.cas_media_agb[0],cas_total[0]),div(self.cas_media_agb[1],cas_total[1]),div(self.cas_media_agb[2],cas_total[2]),div(self.cas_media_agb[0]+self.cas_media_agb[1]+self.cas_media_agb[2],cas_total[0]+cas_total[1]+cas_total[2]),div(self.cas_media_agb[3],cas_total[3]),div(self.cas_media_agb[4],cas_total[4]),div(self.cas_media_agb[5],cas_total[5]),div(self.cas_media_agb[3]+self.cas_media_agb[4]+self.cas_media_agb[5],cas_total[3]+cas_total[4]+cas_total[5]),div(self.cas_media_agb[6],cas_total[6]),div(self.cas_media_agb[7],cas_total[7]),div(self.cas_media_agb[8],cas_total[8]),div(self.cas_media_agb[6]+self.cas_media_agb[7]+self.cas_media_agb[8],cas_total[6]+cas_total[7]+cas_total[8]),div(self.cas_media_agb[9],cas_total[9]),div(self.cas_media_agb[10],cas_total[10]),div(self.cas_media_agb[11],cas_total[11]),div(self.cas_media_agb[9]+self.cas_media_agb[10]+self.cas_media_agb[11],cas_total[9]+cas_total[10]+cas_total[11]),div(sum(self.cas_media_agb.values()),sum(cas_total.values()))])

        data.append(['CAS Media AGB',div(self.cas_media_agb_ly,self.total_customer_ly['DM Pool']),div(self.cas_media_agb[0],cas_total[0]),div(self.cas_media_agb[1],cas_total[1]),div(self.cas_media_agb[2],cas_total[2]),div(self.cas_media_agb[0]+self.cas_media_agb[1]+self.cas_media_agb[2],cas_total[0]+cas_total[1]+cas_total[2]),div(self.cas_media_agb[3],cas_total[3]),div(self.cas_media_agb[4],cas_total[4]),div(self.cas_media_agb[5],cas_total[5]),div(self.cas_media_agb[3]+self.cas_media_agb[4]+self.cas_media_agb[5],cas_total[3]+cas_total[4]+cas_total[5]),div(self.cas_media_agb[6],cas_total[6]),div(self.cas_media_agb[7],cas_total[7]),div(self.cas_media_agb[8],cas_total[8]),div(self.cas_media_agb[6]+self.cas_media_agb[7]+self.cas_media_agb[8],cas_total[6]+cas_total[7]+cas_total[8]),div(self.cas_media_agb[9],cas_total[9]),div(self.cas_media_agb[10],cas_total[10]),div(self.cas_media_agb[11],cas_total[11]),div(self.cas_media_agb[9]+self.cas_media_agb[10]+self.cas_media_agb[11],cas_total[9]+cas_total[10]+cas_total[11]),div(sum(self.cas_media_agb.values()),sum(cas_total.values()))])

        for c in self.fn.cas_dm_pool_channels:
            if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                v2 = self.cas_others_total_for_months['DM Pool']
                last_year2 = self.cas_others_total_customer_ly['DM Pool']
                total_amount = sum(v2.values())
                v1 = self.soi_cas_others
                total_amount1 = sum(v1.values())
                last_year1 = self.soi_cas_others_channels_ly
            else:
                v2 = self.grp_channels["{0}_{1}".format('DM Pool',c)]
                last_year2 = self.fn.channel_total_ly('DM Pool',c)
                total_amount = sum(v2)
                v1 = self.soi_total2[c]
                total_amount1 = sum(v1)
                last_year1 = self.soi_channel_total_ly[c]

            data.append([c,div(last_year1,last_year2),div(v1[0],v2[0]),div(v1[1],v2[1]),div(v1[2],v2[2]),div(v1[0]+v1[1]+v1[2],v2[0]+v2[1]+v2[2]),div(v1[3],v2[3]),div(v1[4],v2[4]),div(v1[5],v2[5]),div(v1[3]+v1[4]+v1[5],v2[3]+v2[4]+v2[5]),div(v1[6],v2[6]),div(v1[7],v2[7]),div(v1[8],v2[8]),div(v1[6]+v1[7]+v1[8],v2[6]+v2[7]+v2[8]),div(v1[9],v2[9]),div(v1[10],v2[10]),div(v1[11],v2[11]),div(v1[9]+v1[10]+v1[11],v2[9]+v2[10]+v2[11]),div(total_amount1,total_amount)])

        return data

    def cpp_18_50_table(self):
        res = self.cpp_4_plus_table()
        res[0] = ['Avg CPP 18-50']
        return res


    def ratio_shr_helper(self):
        values = self.fn.non_shr_channel_value(13,31)
        a2 = {}
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a2[i] = total
        return a2


    def ratio_total_table(self,dm_header):
        data = [['POWER RATIO* TOTAL 4+']]
        data.append(['Row Labels']+self.headers)

        shr_menu = self.ratio_shr_helper()
  
        data.append(['CAS Media',calculate_percente_2(div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]),shr_menu[0]),calculate_percente_2(div_round_1(self.cas_media_total[0],dm_header[1]),shr_menu[1]),calculate_percente_2(div_round_1(self.cas_media_total[1],dm_header[2]),shr_menu[2]),calculate_percente_2(div_round_1(self.cas_media_total[2],dm_header[3]),shr_menu[3]),calculate_percente_2(div_round_1(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2],dm_header[4]),shr_menu[4]),calculate_percente_2(div_round_1(self.cas_media_total[3],dm_header[5]),shr_menu[5]),calculate_percente_2(div_round_1(self.cas_media_total[4],dm_header[6]),shr_menu[6]),calculate_percente_2(div_round_1(self.cas_media_total[5],dm_header[7]),shr_menu[7]),calculate_percente_2(div_round_1(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5],dm_header[8]),shr_menu[8]),calculate_percente_2(div_round_1(self.cas_media_total[6],dm_header[9]),shr_menu[9]),calculate_percente_2(div_round_1(self.cas_media_total[7],dm_header[10]),shr_menu[10]),calculate_percente_2(div_round_1(self.cas_media_total[8],dm_header[11]),shr_menu[11]),calculate_percente_2(div_round_1(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8],dm_header[12]),shr_menu[12]),calculate_percente_2(div_round_1(self.cas_media_total[9],dm_header[13]),shr_menu[13]),calculate_percente_2(div_round_1(self.cas_media_total[10],dm_header[14]),shr_menu[14]),calculate_percente_2(div_round_1(self.cas_media_total[11],dm_header[15]),shr_menu[15]),calculate_percente_2(div_round_1(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11],dm_header[16]),shr_menu[16]),calculate_percente_2(div_round_1(sum(self.cas_media_total.values()),dm_header[17]),shr_menu[17])])


        channels = self.fn.soi_e2_channels()
        channels.insert(-3,'CAS "non AGB" (SK, Nova Sport, Brainz)')
        channels.insert(2,'CAS Others (w/o SK, Nova Sport, Brainz)')

        shr_values = {}
        for d in self.fn.non_shr_channel_value(13,31):
            if d[0] in self.fn.dm_pool_shr_channels:
                shr_values[d[0]] = d[1:]


        # 1 - div_round_1(v[1],self.cas_media_total[1])

        # 2 - shr_values[c][0]
      
        # 3 - div_round_1(self.cas_media_total[0],dm_header[1])

        for c in channels:

            if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                v = self.soi_cas_others
                ly = div_round_1(self.soi_cas_others_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            elif c == 'CAS "non AGB" (SK, Nova Sport, Brainz)':
                v = self.soi_agb_channal
                ly = div_round_1(self.cas_media_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            else:
                v = self.soi_total2[c]
                ly = div_round_1(self.soi_channel_total_ly[c],sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v)



            data.append([c,calculate_percente_3(ly,percente_to_float(shr_values[c][0]),div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0])),
              calculate_percente_3(div_round_1(v[0],self.cas_media_total[0]),percente_to_float(shr_values[c][1]),div_round_1(self.cas_media_total[0],dm_header[1])),
              calculate_percente_3(div_round_1(v[1],self.cas_media_total[1]),percente_to_float(shr_values[c][2]),div_round_1(self.cas_media_total[1],dm_header[2])),
              calculate_percente_3(div_round_1(v[2],self.cas_media_total[2]),percente_to_float(shr_values[c][3]),div_round_1(self.cas_media_total[2],dm_header[3])),

              calculate_percente_3(div_round_1(v[0]+v[1]+v[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),percente_to_float(shr_values[c][4]),div_round_1(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2],dm_header[4])),

              calculate_percente_3(div_round_1(v[3],self.cas_media_total[3]),percente_to_float(shr_values[c][5]),div_round_1(self.cas_media_total[3],dm_header[5])),
              calculate_percente_3(div_round_1(v[4],self.cas_media_total[4]),percente_to_float(shr_values[c][6]),div_round_1(self.cas_media_total[4],dm_header[6])),
              calculate_percente_3(div_round_1(v[5],self.cas_media_total[5]),percente_to_float(shr_values[c][7]),div_round_1(self.cas_media_total[5],dm_header[7])),

              calculate_percente_3(div_round_1(v[3]+v[4]+v[5],self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5]),percente_to_float(shr_values[c][8]),div_round_1(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5],dm_header[8])),

              calculate_percente_3(div_round_1(v[6],self.cas_media_total[6]),percente_to_float(shr_values[c][9]),div_round_1(self.cas_media_total[6],dm_header[9])),
              calculate_percente_3(div_round_1(v[7],self.cas_media_total[7]),percente_to_float(shr_values[c][10]),div_round_1(self.cas_media_total[7],dm_header[10])),
              calculate_percente_3(div_round_1(v[8],self.cas_media_total[8]),percente_to_float(shr_values[c][11]),div_round_1(self.cas_media_total[8],dm_header[11])),

              calculate_percente_3(div_round_1(v[6]+v[7]+v[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),percente_to_float(shr_values[c][12]),div_round_1(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8],dm_header[12])),

              calculate_percente_3(div_round_1(v[9],self.cas_media_total[9]),percente_to_float(shr_values[c][13]),div_round_1(self.cas_media_total[9],dm_header[13])),
              calculate_percente_3(div_round_1(v[10],self.cas_media_total[10]),percente_to_float(shr_values[c][14]),div_round_1(self.cas_media_total[10],dm_header[14])),
              calculate_percente_3(div_round_1(v[11],self.cas_media_total[11]),percente_to_float(shr_values[c][15]),div_round_1(self.cas_media_total[11],dm_header[15])),

              calculate_percente_3(div_round_1(v[9]+v[10]+v[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),percente_to_float(shr_values[c][16]),div_round_1(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11],dm_header[16])),

              calculate_percente_3(div_round_1(total_amount,sum(self.cas_media_total.values())),percente_to_float(shr_values[c][17]),div_round_1(sum(self.cas_media_total.values()),dm_header[17]))
            ])
        return data


    def ratio_shr_helper_2(self):
        values = self.fn.non_shr_channel_value(45,63)
        a2 = {}
        s = [ (f) for (ch,*f) in values if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in range(0,18):
            total = 0
            for p in s:
                total = total + float(p[i][:-1])
            a2[i] = total
        return a2


    def ratio_18_50_table(self,dm_header):
        data = [['POWER RATIO* TOTAL 18-50']]
        data.append(['Row Labels']+self.headers)

        shr_menu = self.ratio_shr_helper_2()

        data.append(['CAS Media',calculate_percente_2(div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]),shr_menu[0]),calculate_percente_2(div_round_1(self.cas_media_total[0],dm_header[1]),shr_menu[1]),calculate_percente_2(div_round_1(self.cas_media_total[1],dm_header[2]),shr_menu[2]),calculate_percente_2(div_round_1(self.cas_media_total[2],dm_header[3]),shr_menu[3]),calculate_percente_2(div_round_1(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2],dm_header[4]),shr_menu[4]),calculate_percente_2(div_round_1(self.cas_media_total[3],dm_header[5]),shr_menu[5]),calculate_percente_2(div_round_1(self.cas_media_total[4],dm_header[6]),shr_menu[6]),calculate_percente_2(div_round_1(self.cas_media_total[5],dm_header[7]),shr_menu[7]),calculate_percente_2(div_round_1(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5],dm_header[8]),shr_menu[8]),calculate_percente_2(div_round_1(self.cas_media_total[6],dm_header[9]),shr_menu[9]),calculate_percente_2(div_round_1(self.cas_media_total[7],dm_header[10]),shr_menu[10]),calculate_percente_2(div_round_1(self.cas_media_total[8],dm_header[11]),shr_menu[11]),calculate_percente_2(div_round_1(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8],dm_header[12]),shr_menu[12]),calculate_percente_2(div_round_1(self.cas_media_total[9],dm_header[13]),shr_menu[13]),calculate_percente_2(div_round_1(self.cas_media_total[10],dm_header[14]),shr_menu[14]),calculate_percente_2(div_round_1(self.cas_media_total[11],dm_header[15]),shr_menu[15]),calculate_percente_2(div_round_1(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11],dm_header[16]),shr_menu[16]),calculate_percente_2(div_round_1(sum(self.cas_media_total.values()),dm_header[17]),shr_menu[17])])


        channels = self.fn.soi_e2_channels()
        channels.insert(-3,'CAS "non AGB" (SK, Nova Sport, Brainz)')
        channels.insert(2,'CAS Others (w/o SK, Nova Sport, Brainz)')

        shr_values = {}
        for d in self.fn.non_shr_channel_value(45,63):
            if d[0] in self.fn.dm_pool_shr_channels:
                shr_values[d[0]] = d[1:]


        # 1 - div_round_1(v[1],self.cas_media_total[1])

        # 2 - shr_values[c][0]
      
        # 3 - div_round_1(self.cas_media_total[0],dm_header[1])

        for c in channels:

            if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                v = self.soi_cas_others
                ly = div_round_1(self.soi_cas_others_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            elif c == 'CAS "non AGB" (SK, Nova Sport, Brainz)':
                v = self.soi_agb_channal
                ly = div_round_1(self.cas_media_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            else:
                v = self.soi_total2[c]
                ly = div_round_1(self.soi_channel_total_ly[c],sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v)

            data.append([c,calculate_percente_3(ly,percente_to_float(shr_values[c][0]),div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0])),
              calculate_percente_3(div_round_1(v[0],self.cas_media_total[0]),percente_to_float(shr_values[c][1]),div_round_1(self.cas_media_total[0],dm_header[1])),
              calculate_percente_3(div_round_1(v[1],self.cas_media_total[1]),percente_to_float(shr_values[c][2]),div_round_1(self.cas_media_total[1],dm_header[2])),
              calculate_percente_3(div_round_1(v[2],self.cas_media_total[2]),percente_to_float(shr_values[c][3]),div_round_1(self.cas_media_total[2],dm_header[3])),

              calculate_percente_3(div_round_1(v[0]+v[1]+v[2],self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2]),percente_to_float(shr_values[c][4]),div_round_1(self.cas_media_total[0]+self.cas_media_total[1]+self.cas_media_total[2],dm_header[4])),

              calculate_percente_3(div_round_1(v[3],self.cas_media_total[3]),percente_to_float(shr_values[c][5]),div_round_1(self.cas_media_total[3],dm_header[5])),
              calculate_percente_3(div_round_1(v[4],self.cas_media_total[4]),percente_to_float(shr_values[c][6]),div_round_1(self.cas_media_total[4],dm_header[6])),
              calculate_percente_3(div_round_1(v[5],self.cas_media_total[5]),percente_to_float(shr_values[c][7]),div_round_1(self.cas_media_total[5],dm_header[7])),

              calculate_percente_3(div_round_1(v[3]+v[4]+v[5],self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5]),percente_to_float(shr_values[c][8]),div_round_1(self.cas_media_total[3]+self.cas_media_total[4]+self.cas_media_total[5],dm_header[8])),

              calculate_percente_3(div_round_1(v[6],self.cas_media_total[6]),percente_to_float(shr_values[c][9]),div_round_1(self.cas_media_total[6],dm_header[9])),
              calculate_percente_3(div_round_1(v[7],self.cas_media_total[7]),percente_to_float(shr_values[c][10]),div_round_1(self.cas_media_total[7],dm_header[10])),
              calculate_percente_3(div_round_1(v[8],self.cas_media_total[8]),percente_to_float(shr_values[c][11]),div_round_1(self.cas_media_total[8],dm_header[11])),

              calculate_percente_3(div_round_1(v[6]+v[7]+v[8],self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8]),percente_to_float(shr_values[c][12]),div_round_1(self.cas_media_total[6]+self.cas_media_total[7]+self.cas_media_total[8],dm_header[12])),

              calculate_percente_3(div_round_1(v[9],self.cas_media_total[9]),percente_to_float(shr_values[c][13]),div_round_1(self.cas_media_total[9],dm_header[13])),
              calculate_percente_3(div_round_1(v[10],self.cas_media_total[10]),percente_to_float(shr_values[c][14]),div_round_1(self.cas_media_total[10],dm_header[14])),
              calculate_percente_3(div_round_1(v[11],self.cas_media_total[11]),percente_to_float(shr_values[c][15]),div_round_1(self.cas_media_total[11],dm_header[15])),

              calculate_percente_3(div_round_1(v[9]+v[10]+v[11],self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11]),percente_to_float(shr_values[c][16]),div_round_1(self.cas_media_total[9]+self.cas_media_total[10]+self.cas_media_total[11],dm_header[16])),

              calculate_percente_3(div_round_1(total_amount,sum(self.cas_media_total.values())),percente_to_float(shr_values[c][17]),div_round_1(sum(self.cas_media_total.values()),dm_header[17]))
            ])
        return data

    #########################################
    ########### MARKO TABLES ################
    #########################################

    def marko_sov1_sov2_table(self,customers,sov1,sov2):
        s = ['Total','SOV Total','18-50','SOV 18-50']
        new_header = []
        for x in s:
            d = [x]
            for i in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))-1):
                d += ['']
            new_header += d
        
        data2 = [['']+new_header]

        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True)]
        header_list = [item for hl in header_list for item in hl]

        data2.append(['Row Labels']+header_list)

        data2.append(['Grand Total'] + sov1['grand'] + sov2['grand'])
        data2.append(['Nova S'] + sov1['nova_s'] + sov2['nova_s'])
        data2.append(['N1'] + sov1['n1'] + sov2['n1'])
        data2.append(['CAS Others (w/o SK, Nova Sport, Brainz)'] + sov1['cas_others']+sov2['cas_others'])

        for cas in customers:
            data2.append([cas]+sov1["{0}_total".format(cas)] + sov2["{0}_total".format(cas)])
            for c in ['N1','Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']:
                data2.append([c]+sov1["{0}_{1}".format(cas,c)] + sov2["{0}_{1}".format(cas,c)])

        return data2

    def marko_duration(self,cover_duration,customers):
        data2 = []
        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True)]
        header_list = [item for hl in header_list for item in hl]

        data2.append(['']+header_list)

        n1 = self.n1_nova_s_duration['N1']
        nova_s = self.n1_nova_s_duration['Nova S']

        # Cas Media
        cas_media = ['Cas Media','{0} %'.format(round(self.duration_ly['CAS Media'],2))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            cas_media += [calculate_percente(self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3],self.duration_total_by_month[i]+self.duration_total_by_month[i+1]+self.duration_total_by_month[i+2])]

        cas_media += [calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1],self.duration_total_by_month[self.fn.range_by_quarter[self.QUARTER]]),calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2],self.duration_total_by_month[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+3],self.duration_total_by_month[self.fn.range_by_quarter[self.QUARTER]+2])]

        i = self.fn.range_by_quarter[self.QUARTER]
        cas_media += [calculate_percente(self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3],self.duration_total_by_month[i]+self.duration_total_by_month[i+1]+self.duration_total_by_month[i+2])]
        
        cas_media += [calculate_percente(self.fn.sum_delta_time(self.all_duration.values()),self.fn.sum_delta_time(self.duration_total_by_month))]

        data2.append(cas_media)

        # N1
        n = ['N1','{0} %'.format(round(self.duration_ly['N1'],2))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            n += [calculate_percente(n1[i]+n1[i+1]+n1[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        n += [calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]+1],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(n1[self.fn.range_by_quarter[self.QUARTER]+2],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+3])]

        i = self.fn.range_by_quarter[self.QUARTER]
        n += [calculate_percente(n1[i]+n1[i+1]+n1[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        
        n += [calculate_percente(self.fn.sum_delta_time(n1.values()),self.fn.sum_delta_time(self.all_duration.values()))]
        data2.append(n)

        # Nova S
        nova=['Nova S','{0} %'.format(round(self.duration_ly['Nova S'],2))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            nova += [calculate_percente(nova_s[i]+nova_s[i+1]+nova_s[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        nova += [calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]+1],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(nova_s[self.fn.range_by_quarter[self.QUARTER]+2],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+3])]

        i = self.fn.range_by_quarter[self.QUARTER]
        nova += [calculate_percente(nova_s[i]+nova_s[i+1]+nova_s[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        
        nova += [calculate_percente(self.fn.sum_delta_time(nova_s.values()),self.fn.sum_delta_time(self.all_duration.values()))]
        data2.append(nova)

        # Cas Others
        cas_others = ['CAS Others (w/o SK, Nova Sport, Brainz)','{0} %'.format(round(self.duration_ly['CAS Others (w/o SK, Nova Sport, Brainz)'],2))]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            cas_others += [calculate_percente(self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3]-nova_s[i]-nova_s[i+1]-nova_s[i+2]-n1[i]-n1[i+1]-n1[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        cas_others += [calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]-nova_s[self.fn.range_by_quarter[self.QUARTER]]-n1[self.fn.range_by_quarter[self.QUARTER]],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]-nova_s[self.fn.range_by_quarter[self.QUARTER]+1]-n1[self.fn.range_by_quarter[self.QUARTER]+1],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(self.all_duration[self.fn.range_by_quarter[self.QUARTER]+3]-nova_s[self.fn.range_by_quarter[self.QUARTER]+2]-n1[self.fn.range_by_quarter[self.QUARTER]+2],self.all_duration[self.fn.range_by_quarter[self.QUARTER]+3])]

        i = self.fn.range_by_quarter[self.QUARTER]
        cas_others += [calculate_percente(self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3]-nova_s[i]-nova_s[i+1]-nova_s[i+2]-n1[i]-n1[i+1]-n1[i+2],self.all_duration[i+1]+self.all_duration[i+2]+self.all_duration[i+3])]
        
        cas_others += [calculate_percente(self.fn.sum_delta_time(self.all_duration.values())-self.fn.sum_delta_time(n1.values())-self.fn.sum_delta_time(nova_s.values()),self.fn.sum_delta_time(self.all_duration.values()))]
        data2.append(cas_others)

        for cas in customers:
            cas_total = cover_duration[cas]
            cas_customer_total = self.cas_duration[cas]

            data2.append([cas])
            c = ['Cas Media',cas_total[0]]
            for q in range(1,self.QUARTER):
                i = self.fn.duration_by_quarter[q]
                c += [cas_total[i+3]]

            i = self.fn.duration_by_quarter[self.QUARTER]
            c += [cas_total[self.fn.duration_by_quarter[self.QUARTER]],cas_total[self.fn.duration_by_quarter[self.QUARTER]+1],cas_total[self.fn.duration_by_quarter[self.QUARTER]+2],cas_total[i+3],cas_total[17]]

            data2.append(c)

            for c in ['N1','Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']:
                if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                    value = self.cas_others_duration[cas]
                    total_amount = self.fn.sum_delta_time(value.values())
                else:
                    value = self.c_duration["{0}_{1}".format(cas,c)]
                    total_amount = self.fn.sum_delta_time(value)
                
                ch = [c,'{0} %'.format(round(self.duration_ly["{0}_{1}".format(cas,c)],2))]
                for q in range(1,self.QUARTER):
                    i = self.fn.range_by_quarter[q]
                    ch += [calculate_percente(value[i]+value[i+1]+value[i+2],cas_customer_total[i]+cas_customer_total[i+1]+cas_customer_total[i+2])]
                
                i = self.fn.range_by_quarter[self.QUARTER]
                ch += [calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]],cas_customer_total[self.fn.range_by_quarter[self.QUARTER]]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+1],cas_customer_total[self.fn.range_by_quarter[self.QUARTER]+1]),calculate_percente(value[self.fn.range_by_quarter[self.QUARTER]+2],cas_customer_total[self.fn.range_by_quarter[self.QUARTER]+2]),calculate_percente(value[i]+value[i+1]+value[i+2],cas_customer_total[i]+cas_customer_total[i+1]+cas_customer_total[i+2]),calculate_percente(total_amount,self.fn.sum_delta_time(cas_customer_total.values()))]

                data2.append(ch)
        return data2


    def marko_shr_table(self):
        new_data = ['Total', "SHR% of Audience TOTAL 4+ (08-24 h)"]
        
        for i in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))-1):
            new_data += ['']
        
        new_data += ["SHR% of Audience 18-50 (08-24 h)"]

        data = [new_data]

        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True)]
        header_list = [item for hl in header_list for item in hl]

        data.append(['Row Labels']+header_list)

        values1 = self.fn.non_shr_channel_value(13,31)

        a1 = ['CAS Media']
        s1 = [ (f) for (ch,*f) in values1 if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']]

        for i in self.fn.marko_shr_total[self.QUARTER]:
            total = 0
            for p in s1:
                total = total + float(p[i][:-1])
            a1 += ['{0}%'.format(round(total,2))]

        a2 = ['CAS Media AGB']
        s2 = [ (f) for (ch,*f) in values1 if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in self.fn.marko_shr_total[self.QUARTER]:
            total = 0
            for p in s2:
                total = total + float(p[i][:-1])
            a2 += ['{0}%'.format(round(total,2))]

        values2 = self.fn.non_shr_channel_value(45,63)

        s3 = [ (f) for (ch,*f) in values2 if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)']]
        for i in self.fn.marko_shr_total[self.QUARTER]:
            total = 0
            for p in s3:
                total = total + float(p[i][:-1])
            a1 += ['{0}%'.format(round(total,2))]

        s4 = [ (f) for (ch,*f) in values2 if ch in ['N1', 'Nova S','CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)']]
        for i in self.fn.marko_shr_total[self.QUARTER]:
            total = 0
            for p in s4:
                total = total + float(p[i][:-1])
            a2 += ['{0}%'.format(round(total,2))]

        data.append(a1)
        data.append(a2)

        val1 = self.fn.non_shr_channel_value_2(13,31)
        val2 = self.fn.non_shr_channel_value_2(45,63)

        res1 = []
        for d in val1:
            n = []
            if d[0] in self.fn.dm_pool_shr_channels and d[0] not in ['CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)','Brainz']:
                n += [d[0]]
                for i in self.fn.marko_shr_total[self.QUARTER]:
                    n += [d[1:][i]]
            if len(n) > 0:
                res1.append(n)

        res2 = []
        for d in val2:
            n = []
            if d[0] in self.fn.dm_pool_shr_channels and d[0] not in ['CAS Others (w/o SK, Nova Sport, Brainz)','CAS "non AGB" (SK, Nova Sport, Brainz)','Brainz']:
                for i in self.fn.marko_shr_total[self.QUARTER]:
                    n += [d[1:][i]]
            if len(n) > 0:
                res2.append(n)

        for s in range(len(res1)):
            data.append(res1[s] + res2[s])
        return data

    def marko_soi_dm(self, soi_header):
        data = [['', "COST, Net Eur; SO%"]]

        new_header = []
        for h in get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True):
            new_header.append(h)
            new_header.append('')

        data.append(['Row Labels']+new_header)

        #DM Pool
        dm_pool = ['DM Pool']
        for i in self.fn.marko_shr_total[self.QUARTER]:
            dm_pool += [float_with_comma(round(soi_header[i])), '']

        data.append(dm_pool)

        #CAS MEDIA
        cas_media = ['CAS Media',float_with_comma(round(sum(self.soi_channel_total_ly.values()))),calculate_percente_round_1(sum(self.soi_channel_total_ly.values()),soi_header[0])]

        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            cas_media += [float_with_comma(round(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2])),calculate_percente_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],soi_header[q*4])]

        cas_media += [float_with_comma(round(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]])),calculate_percente_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]],soi_header[self.fn.duration_by_quarter[self.QUARTER]]),float_with_comma(round(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1])),calculate_percente_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1],soi_header[self.fn.duration_by_quarter[self.QUARTER]+1]),float_with_comma(round(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2])),calculate_percente_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2],soi_header[self.fn.duration_by_quarter[self.QUARTER]+2])]
        
        i = self.fn.range_by_quarter[self.QUARTER]
        cas_media += [float_with_comma(round(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2])),calculate_percente_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],soi_header[self.QUARTER*4])]

        cas_media += [float_with_comma(round(sum(self.cas_media_total.values()))),calculate_percente_round_1(sum(self.cas_media_total.values()),soi_header[17])]

        data.append(cas_media)

        #CAS MEDIA AGB
        agb = ['CAS Media AGB',float_with_comma(round(self.cas_media_agb_ly)),calculate_percente_round_1(self.cas_media_agb_ly,sum(self.soi_channel_total_ly.values()))]
    
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            agb += [float_with_comma(round(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2])),calculate_percente_round_1(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2],self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2])]

        agb += [float_with_comma(round(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]])),calculate_percente_round_1(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]]),float_with_comma(round(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+1])),calculate_percente_round_1(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+1],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1]),float_with_comma(round(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+2])),calculate_percente_round_1(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+2],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2])]

        i = self.fn.range_by_quarter[self.QUARTER]

        agb += [float_with_comma(round(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2])),calculate_percente_round_1(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2],self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2])]
        
        agb += [float_with_comma(round(sum(self.cas_media_agb.values()))),calculate_percente_round_1(sum(self.cas_media_agb.values()),sum(self.cas_media_total.values()))]

        data.append(agb)

        channels = self.fn.soi_e2_channels()

        for c in channels:
            value = self.soi_total2[c]

            channel_row = [c,float_with_comma(round(self.soi_channel_total_ly[c])),calculate_percente_round_1(self.soi_channel_total_ly[c],sum(self.soi_channel_total_ly.values()))]

            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                channel_row += [float_with_comma(round(value[i]+value[i+1]+value[i+2])),calculate_percente_round_1(value[i]+value[i+1]+value[i+2],self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2])]

            channel_row += [float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]]),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+1])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]+1],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1]),float_with_comma(round(value[self.fn.range_by_quarter[self.QUARTER]+2])),calculate_percente_round_1(value[self.fn.range_by_quarter[self.QUARTER]+2],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2])]
            
            i2 = self.fn.range_by_quarter[self.QUARTER]
            channel_row += [float_with_comma(round(value[i2]+value[i2+1]+value[i2+2])),calculate_percente_round_1(value[i2]+value[i2+1]+value[i2+2],self.cas_media_total[i2]+self.cas_media_total[i2+1]+self.cas_media_total[i2+2])]

            channel_row += [float_with_comma(round(sum(value))),calculate_percente_round_1(sum(value),sum(self.cas_media_total.values()))]

            data.append(channel_row)

        return data
            

    def marko_average_cpp(self):
        cas_total = self.grp_customers_v['DM Pool']
        res = []
        data = ['CAS Media AGB',div(self.cas_media_agb_ly,self.total_customer_ly['DM Pool'])]
        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            data += [div(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2],cas_total[i]+cas_total[i+1]+cas_total[i+2])]
        
        i = self.fn.range_by_quarter[self.QUARTER]
        data += [div(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]],cas_total[self.fn.range_by_quarter[self.QUARTER]]),div(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+1],cas_total[self.fn.range_by_quarter[self.QUARTER]+1]),div(self.cas_media_agb[self.fn.range_by_quarter[self.QUARTER]+2],cas_total[self.fn.range_by_quarter[self.QUARTER]+2]),div(self.cas_media_agb[i]+self.cas_media_agb[i+1]+self.cas_media_agb[i+2],cas_total[i]+cas_total[i+1]+cas_total[i+2]),div(sum(self.cas_media_agb.values()),sum(cas_total.values()))]

        res.append(data)

        for c in self.fn.cas_dm_pool_channels:
            if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
               continue

            v2 = self.grp_channels["{0}_{1}".format('DM Pool',c)]
            last_year2 = self.fn.channel_total_ly('DM Pool',c)
            total_amount = sum(v2)
            v1 = self.soi_total2[c]
            total_amount1 = sum(v1)
            last_year1 = self.soi_channel_total_ly[c]

            new_arr = [c,div(last_year1,last_year2)]
            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                new_arr += [div(v1[i]+v1[i+1]+v1[i+2],v2[i]+v2[i+1]+v2[i+2])]
            
            i = self.fn.range_by_quarter[self.QUARTER]
            new_arr += [div(v1[self.fn.range_by_quarter[self.QUARTER]],v2[self.fn.range_by_quarter[self.QUARTER]]),div(v1[self.fn.range_by_quarter[self.QUARTER]+1],v2[self.fn.range_by_quarter[self.QUARTER]+1]),div(v1[self.fn.range_by_quarter[self.QUARTER]+2],v2[self.fn.range_by_quarter[self.QUARTER]+2]),div(v1[i]+v1[i+1]+v1[i+2],v2[i]+v2[i+1]+v2[i+2]),div(total_amount1,total_amount)]

            res.append(new_arr)

        return res

    def merge_tables(self,table1,table2):
        new_data = ["Avg CPP 4+"]
        
        for i in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))-1):
            new_data += ['']
        
        new_data += ["Avg Cpp 18-50"]

        data = [['']+new_data]
        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True)]
        header_list = [item for hl in header_list for item in hl]

        data.append(['Row Labels']+header_list)
        for s in range(len(table1)):
            res = table1[s] + table2[s][1:]
            data.append(res)
        return data

    def marko_ratio_total(self,dm_header):
        new_data = ["POWER RATIO* TOTAL 4+"]
        
        for i in range(len(get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True))-1):
            new_data += ['']
        
        new_data += ["POWER RATIO* TOTAL 18-50"]

        data = [['']+new_data]

        header_list = ['',get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True),get_quarter_header(self.CURRENT_YEAR,self.QUARTER,True)]
        header_list = [item for hl in header_list for item in hl]

        data.append(['Row Labels']+header_list)

        shr_menu = self.ratio_shr_helper()
        shr_menu_2 = self.ratio_shr_helper_2()
        ii = self.fn.range_by_quarter[self.QUARTER]

        agb = ['CAS Media',calculate_percente_2(div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]),shr_menu[0])]

        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            agb += [calculate_percente_2(div_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],dm_header[q*4]),shr_menu[q*4])]

        agb += [calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]],dm_header[self.fn.duration_by_quarter[self.QUARTER]]),shr_menu[self.fn.duration_by_quarter[self.QUARTER]]),calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1],dm_header[self.fn.duration_by_quarter[self.QUARTER]+1]),shr_menu[self.fn.duration_by_quarter[self.QUARTER]+1]),calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2],dm_header[self.fn.duration_by_quarter[self.QUARTER]+2]),shr_menu[self.fn.duration_by_quarter[self.QUARTER]+2])]
        
        agb += [calculate_percente_2(div_round_1(self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2],dm_header[self.QUARTER*4]),shr_menu[self.QUARTER*4])]
        
        agb += [calculate_percente_2(div_round_1(sum(self.cas_media_total.values()),dm_header[17]),shr_menu[17])]

        agb += [calculate_percente_2(div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]),shr_menu_2[0])]

        for q in range(1,self.QUARTER):
            i = self.fn.range_by_quarter[q]
            agb += [calculate_percente_2(div_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],dm_header[q*4]),shr_menu_2[q*4])]

        agb += [calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]],dm_header[self.fn.duration_by_quarter[self.QUARTER]]),shr_menu_2[self.fn.duration_by_quarter[self.QUARTER]]),calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1],dm_header[self.fn.duration_by_quarter[self.QUARTER]+1]),shr_menu_2[self.fn.duration_by_quarter[self.QUARTER]+1]),calculate_percente_2(div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2],dm_header[self.fn.duration_by_quarter[self.QUARTER]+2]),shr_menu_2[self.fn.duration_by_quarter[self.QUARTER]+2])]

        agb += [calculate_percente_2(div_round_1(self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2],dm_header[self.QUARTER*4]),shr_menu_2[self.QUARTER*4])]
        
        agb += [calculate_percente_2(div_round_1(sum(self.cas_media_total.values()),dm_header[17]),shr_menu_2[17])]

        data.append(agb)

        channels = self.fn.soi_e2_channels()
        channels.insert(-3,'CAS "non AGB" (SK, Nova Sport, Brainz)')
        channels.insert(2,'CAS Others (w/o SK, Nova Sport, Brainz)')

        shr_values = {}
        for d in self.fn.non_shr_channel_value(13,31):
            if d[0] in self.fn.dm_pool_shr_channels:
                shr_values[d[0]] = d[1:]

        shr_values_2 = {}
        for d in self.fn.non_shr_channel_value(45,63):
            if d[0] in self.fn.dm_pool_shr_channels:
                shr_values_2[d[0]] = d[1:]


        # 1 - div_round_1(v[1],self.cas_media_total[1])

        # 2 - shr_values[c][0]
      
        # 3 - div_round_1(self.cas_media_total[0],dm_header[1])

        for c in channels:
            if c == 'CAS Others (w/o SK, Nova Sport, Brainz)':
                v = self.soi_cas_others
                ly = div_round_1(self.soi_cas_others_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            elif c == 'CAS "non AGB" (SK, Nova Sport, Brainz)':
                v = self.soi_agb_channal
                ly = div_round_1(self.cas_media_channels_ly,sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v.values())
            else:
                v = self.soi_total2[c]
                ly = div_round_1(self.soi_channel_total_ly[c],sum(self.soi_channel_total_ly.values()))
                total_amount = sum(v)

            new_arr = [c,calculate_percente_3(ly,percente_to_float(shr_values[c][0]),div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]))]

            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                new_arr += [calculate_percente_3(div_round_1(v[i]+v[i+1]+v[i+2],self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2]),percente_to_float(shr_values[c][q*4]),div_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],dm_header[q*4]))]

            new_arr += [calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]]),percente_to_float(shr_values[c][self.fn.duration_by_quarter[self.QUARTER]]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]],dm_header[self.fn.duration_by_quarter[self.QUARTER]])),
            calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]+1],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1]),percente_to_float(shr_values[c][self.fn.duration_by_quarter[self.QUARTER]+1]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1],dm_header[self.fn.duration_by_quarter[self.QUARTER]+1])),
            calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]+2],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2]),percente_to_float(shr_values[c][self.fn.duration_by_quarter[self.QUARTER]+2]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2],dm_header[self.fn.duration_by_quarter[self.QUARTER]+2]))]

            new_arr += [calculate_percente_3(div_round_1(v[ii]+v[ii+1]+v[ii+2],self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2]),percente_to_float(shr_values[c][self.QUARTER*4]),div_round_1(self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2],dm_header[self.QUARTER*4]))]

            new_arr += [calculate_percente_3(div_round_1(total_amount,sum(self.cas_media_total.values())),percente_to_float(shr_values[c][17]),div_round_1(sum(self.cas_media_total.values()),dm_header[17]))]

            new_arr += [calculate_percente_3(ly,percente_to_float(shr_values_2[c][0]),div_round_1(sum(self.soi_channel_total_ly.values()),dm_header[0]))]

            for q in range(1,self.QUARTER):
                i = self.fn.range_by_quarter[q]
                new_arr += [calculate_percente_3(div_round_1(v[i]+v[i+1]+v[i+2],self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2]),percente_to_float(shr_values_2[c][q*4]),div_round_1(self.cas_media_total[i]+self.cas_media_total[i+1]+self.cas_media_total[i+2],dm_header[q*4]))]

            new_arr += [calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]]),percente_to_float(shr_values_2[c][self.fn.duration_by_quarter[self.QUARTER]]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]],dm_header[self.fn.duration_by_quarter[self.QUARTER]])),
            calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]+1],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1]),percente_to_float(shr_values_2[c][self.fn.duration_by_quarter[self.QUARTER]+1]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+1],dm_header[self.fn.duration_by_quarter[self.QUARTER]+1])),
            calculate_percente_3(div_round_1(v[self.fn.range_by_quarter[self.QUARTER]+2],self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2]),percente_to_float(shr_values_2[c][self.fn.duration_by_quarter[self.QUARTER]+2]),div_round_1(self.cas_media_total[self.fn.range_by_quarter[self.QUARTER]+2],dm_header[self.fn.duration_by_quarter[self.QUARTER]+2]))]

            new_arr += [calculate_percente_3(div_round_1(v[ii]+v[ii+1]+v[ii+2],self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2]),percente_to_float(shr_values_2[c][self.QUARTER*4]),div_round_1(self.cas_media_total[ii]+self.cas_media_total[ii+1]+self.cas_media_total[ii+2],dm_header[self.QUARTER*4]))]

            new_arr += [calculate_percente_3(div_round_1(total_amount,sum(self.cas_media_total.values())),percente_to_float(shr_values_2[c][17]),div_round_1(sum(self.cas_media_total.values()),dm_header[17]))]

            data.append(new_arr)
 
        return data


