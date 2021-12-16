from openpyxl import Workbook
from openpyxl import load_workbook
import datetime

class Config():
    now=datetime.datetime.now()
    csvName = 'Kvartalni_TV_'+str('{:02d}_'.format(now.day))+str('{:02d}_'.format(now.month))+str(now.year)+'.xlsx'
    PATH = "Q3_poslednji.xlsx"
    NEW_FILE_PATH = csvName

    wb = Workbook()

    ws = wb.active
    ws.title = 'COVER'
    ws2 = wb.create_sheet('COVER Marko')
    ws3 = wb.create_sheet('COVER Cas')
    ws4 = wb.create_sheet('COVER Cas Marko')

    file = load_workbook(filename = PATH)

    grp_baza = file['BAZA Eq GRP i DUR']
    e2_baza = file['BAZA E2']
    shr_baza = file['BAZA Shr TV']
    cas_baza = file['BAZA CAS']

    channel_map = {
        'RTS': ('RTS 1', 'RTS 2'),
        'B92': ('B92',),
        'PINK': ('PINK',),
        'PINK Cable': ('Pink cable total',),
        'Ostali': ('Other', 'RTS 3', 'RTS Drama', 'RTS Trezor', 'RTS Zivot', 'RTS Muzika', 'RTS Kolo', 'RTS Poletarac', 'RTS Nauka', 'RTS Klasika', 'RTV 1', 'RTV 2', 'Melos Kraljevo', 'Minimax', 'Agro TV', 'Kitchen TV', 'SUPERSTAR 2', 'SUPERSTAR TV', 'FILM KLUB', 'ArenaSport 1', 'Arena Sport 2', 'Arena Sport 3', 'Arena Sport 4', 'Arena Sport 5', 'Toxic TV', 'Dox TV', 'Klasik TV', 'Dexy TV', 'K1', 'Kazbuka', 'TV Doktor', 'Arena Fight', 'Balkan Trip', 'Kurir TV', 'Studio B'),
        'PRVA': ('PRVA',),
        'PRVA Cable': ('PRVA Plus', 'PRVA Max', 'PRVA World', 'PRVA Kick', 'PRVA Life', 'PRVA Files'),
        'Cas Media': ('N1', 'Grand Kanal', 'Cinemania', 'Nova S', 'IDJ', 'SK 1', 'SK 2', 'SK 3', 'Pikaboo', 'Vavoom', 'CineStar', 'CineStar Action', 'Discovery', 'Discovery ID', 'HGTV', 'TLC', 'DIVA - Universal', 'AXN', 'Nickelodeon', 'TV 1000', 'NickJr', 'Nova Sport', 'Animal Planet'),
        'Fox': ('Fox', 'Fox Crime', 'Fox Life', 'Fox Movies', 'National Geographic', '24 Kitchen'),
        'Happy': ('Happy',)
    }

    duration_last_year = {
        'DM Pool': '25.77%',
        'DM Pool_B92': '4.44%',
        'DM Pool_Cas Media': '40.58%',
        'DM Pool_Fox': '20.91%',
        'DM Pool_Happy': '0.85%',
        'DM Pool_Ostali': '0.72%',
        'DM Pool_PINK': '2.92%',
        'DM Pool_PINK Cable': '15.23%',
        'DM Pool_PRVA': '3.52%',
        'DM Pool_PRVA Cable': '7.05%',
        'DM Pool_RTS': '3.77%',
        'MEDIA House': '8.18%',
        'MEDIA House_B92': '1.63%',
        'MEDIA House_Cas Media': '29.77%',
        'MEDIA House_Fox': '27.29%',
        'MEDIA House_Happy': '0.49%',
        'MEDIA House_Ostali': '0.93%',
        'MEDIA House_PINK': '2.15%',
        'MEDIA House_PINK Cable': '29.89%',
        'MEDIA House_PRVA': '2.31%',
        'MEDIA House_PRVA Cable': '2.50%',
        'MEDIA House_RTS': '3.04%',
        'MEDIA Pool': '9.97%',
        'MEDIA Pool_B92': '6.80%',
        'MEDIA Pool_Cas Media': '20.91%',
        'MEDIA Pool_Fox': '15.00%',
        'MEDIA Pool_Happy': '0.39%',
        'MEDIA Pool_Ostali': '1.36%',
        'MEDIA Pool_PINK': '5.69%',
        'MEDIA Pool_PINK Cable': '24.07%',
        'MEDIA Pool_PRVA Cable': '10.87%',
        'MEDIA Pool_PRVA': '7.08%',
        'MEDIA Pool_RTS': '7.83%',
        'MEDIACOM': '12.04%',
        'MEDIACOM_B92': '3.20%',
        'MEDIACOM_Cas Media': '38.43%',
        'MEDIACOM_Fox': '22.29%',
        'MEDIACOM_Happy': '0.66%',
        'MEDIACOM_Ostali': '0.11%',
        'MEDIACOM_PINK': '1.88%',
        'MEDIACOM_PINK Cable': '23.74%',
        'MEDIACOM_PRVA': '2.42%',
        'MEDIACOM_PRVA Cable': '7.07%',
        'MEDIACOM_RTS': '0.19%',
        'ONE MEDIA': '6.62%',
        'ONE MEDIA_B92': '2.27%',
        'ONE MEDIA_Cas Media': '22.69%',
        'ONE MEDIA_Fox': '15.44%',
        'ONE MEDIA_Happy': '',
        'ONE MEDIA_Ostali': '0.04%',
        'ONE MEDIA_PINK': '2.49%',
        'ONE MEDIA_PINK Cable': '44.58%',
        'ONE MEDIA_PRVA': '1.94%',
        'ONE MEDIA_PRVA Cable': '8.60%',
        'ONE MEDIA_RTS': '1.96%',
        'OSTALO (A, N, O)': '37.42%',
        'OSTALO (A, N, O)_B92': '1.98%',
        'OSTALO (A, N, O)_Cas Media': '15.21%',
        'OSTALO (A, N, O)_Fox': '1.95%',
        'OSTALO (A, N, O)_Happy': '4.09%',
        'OSTALO (A, N, O)_Ostali': '14.37%',
        'OSTALO (A, N, O)_PINK': '6.11%',
        'OSTALO (A, N, O)_PINK Cable': '43.81%',
        'OSTALO (A, N, O)_PRVA': '3.38%',
        'OSTALO (A, N, O)_PRVA Cable': '7.42%',
        'OSTALO (A, N, O)_RTS': '1.68%'
    }

    cas_duration_last_year = {
        'CAS Media': 28.45,
        'Nova S': 11.7063002076022,
        'N1': 10.015396722986,
        'CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 11.7063002076022 - 10.015396722986,
        'DM Pool_N1': 8.39270414895303,
        'DM Pool_Nova S': 12.089638389513,
        'DM Pool_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 8.39270414895303 - 12.089638389513,
        'DM Pool_Grand Kanal': 4.30212687282825,
        'DM Pool_Cinemania': 9.0656032014542,
        'DM Pool_IDJ': 0.182893729657949,
        'DM Pool_Pikaboo': 2.07760947318767,
        'DM Pool_Vavoom': 0.152173476125825,
        'DM Pool_CineStar': 5.49154997502275,
        'DM Pool_CineStar Action': 3.21043018995175,
        'DM Pool_Discovery': 8.74937920624431,
        'DM Pool_Discovery ID': 4.50322721810825,
        'DM Pool_HGTV': 0,
        'DM Pool_TLC': 13.4999417515619,
        'DM Pool_DIVA - Universal': 8.96001028737443,
        'DM Pool_AXN': 12.6133923390163,
        'DM Pool_Nickelodeon': 2.13276587865609,
        'DM Pool_TV 1000': 2.45181817429479,
        'DM Pool_Animal Planet': 1.76698751355379,
        'DM Pool_NickJr': 0.357748174495769,
        'DM Pool_SK 1': 0,
        'DM Pool_SK 2': 0,
        'DM Pool_SK 3': 0,
        'DM Pool_Nova Sport': 0,
        'MEDIA House_N1': 5.80371335794949,
        'MEDIA House_Nova S': 12.2906461467799,
        'MEDIA House_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 5.80371335794949 - 12.2906461467799,
        'MEDIA House_Grand Kanal': 0.329170145970634,
        'MEDIA House_Cinemania': 7.37168496214231,
        'MEDIA House_IDJ': 0,
        'MEDIA House_Pikaboo': 4.45692003290139,
        'MEDIA House_Vavoom': 0.0174583596640809,
        'MEDIA House_CineStar': 9.65568365296063,
        'MEDIA House_CineStar Action': 6.5343086059949,
        'MEDIA House_Discovery': 4.24175649064318,
        'MEDIA House_Discovery ID': 9.07897193305051,
        'MEDIA House_HGTV': 0,
        'MEDIA House_TLC': 15.8121183668974,
        'MEDIA House_DIVA - Universal': 9.1846594526277,
        'MEDIA House_AXN': 9.73621863646696,
        'MEDIA House_Nickelodeon': 3.40313031903883,
        'MEDIA House_TV 1000': 1.39799670204946,
        'MEDIA House_Animal Planet': 0.685562834862665,
        'MEDIA House_NickJr': 0,
        'MEDIA House_SK 1': 0,
        'MEDIA House_SK 2': 0,
        'MEDIA House_SK 3': 0,
        'MEDIA House_Nova Sport': 0,
        'MEDIA Pool_N1': 4.34016845986074,
        'MEDIA Pool_Nova S': 12.82448599666,
        'MEDIA Pool_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 4.34016845986074 - 12.82448599666,
        'MEDIA Pool_Grand Kanal': 5.17229264197261,
        'MEDIA Pool_Cinemania': 3.96555972294467,
        'MEDIA Pool_IDJ': 0,
        'MEDIA Pool_Pikaboo': 4.5539372701472,
        'MEDIA Pool_Vavoom': 1.8500926255464,
        'MEDIA Pool_CineStar': 3.35943274837791,
        'MEDIA Pool_CineStar Action': 0.706737481862732,
        'MEDIA Pool_Discovery': 8.95022859802338,
        'MEDIA Pool_Discovery ID': 2.42601363375038,
        'MEDIA Pool_HGTV': 0,
        'MEDIA Pool_TLC': 20.1038501200026,
        'MEDIA Pool_DIVA - Universal': 6.9307635447751,
        'MEDIA Pool_AXN': 15.7989523731304,
        'MEDIA Pool_Nickelodeon': 7.04624889351256,
        'MEDIA Pool_TV 1000': 1.44860879167009,
        'MEDIA Pool_Animal Planet': 0.0820853980160794,
        'MEDIA Pool_NickJr': 0.440541699747219,
        'MEDIA Pool_SK 1': 0,
        'MEDIA Pool_SK 2': 0,
        'MEDIA Pool_SK 3': 0,
        'MEDIA Pool_Nova Sport': 0,
        'MEDIACOM_N1': 5.33803388202536,
        'MEDIACOM_Nova S': 6.20300145165213,
        'MEDIACOM_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 6.20300145165213 - 5.33803388202536,
        'MEDIACOM_Grand Kanal': 4.8403275170622,
        'MEDIACOM_Cinemania': 14.1511926223753,
        'MEDIACOM_IDJ': 0,
        'MEDIACOM_Pikaboo': 2.08163338310603,
        'MEDIACOM_Vavoom': 0,
        'MEDIACOM_CineStar': 5.95325396428907,
        'MEDIACOM_CineStar Action': 3.59726429052943,
        'MEDIACOM_Discovery': 8.45576572748208,
        'MEDIACOM_Discovery ID': 8.40229246318907,
        'MEDIACOM_HGTV': 0,
        'MEDIACOM_TLC': 9.12487024869693,
        'MEDIACOM_DIVA - Universal': 10.0299479062267,
        'MEDIACOM_AXN': 12.7675898976854,
        'MEDIACOM_Nickelodeon': 0.730513456410243,
        'MEDIACOM_TV 1000': 5.65541960065484,
        'MEDIACOM_Animal Planet': 2.51525817759579,
        'MEDIACOM_NickJr': 0.15363541101944,
        'MEDIACOM_SK 1': 0,
        'MEDIACOM_SK 2': 0,
        'MEDIACOM_SK 3': 0,
        'MEDIACOM_Nova Sport': 0,
        'ONE MEDIA_N1': 7.32247088789546,
        'ONE MEDIA_Nova S': 8.61506983994827,
        'ONE MEDIA_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 7.32247088789546 - 8.61506983994827,
        'ONE MEDIA_Grand Kanal': 0,
        'ONE MEDIA_Cinemania': 9.31872044661126,
        'ONE MEDIA_IDJ': 0,
        'ONE MEDIA_Pikaboo': 0.719340711900565,
        'ONE MEDIA_Vavoom': 0.541435163803433,
        'ONE MEDIA_CineStar': 9.9115280517571,
        'ONE MEDIA_CineStar Action': 8.64246425756067,
        'ONE MEDIA_Discovery': 8.56939138840982,
        'ONE MEDIA_Discovery ID': 6.58345427789454,
        'ONE MEDIA_HGTV': 0,
        'ONE MEDIA_TLC': 7.22946905443075,
        'ONE MEDIA_DIVA - Universal': 8.63000074654533,
        'ONE MEDIA_AXN': 9.02940249801659,
        'ONE MEDIA_Nickelodeon': 0.821579462158376,
        'ONE MEDIA_TV 1000': 7.20226443647848,
        'ONE MEDIA_Animal Planet': 6.86340877658935,
        'ONE MEDIA_NickJr': 0,
        'ONE MEDIA_SK 1': 0,
        'ONE MEDIA_SK 2': 0,
        'ONE MEDIA_SK 3': 0,
        'ONE MEDIA_Nova Sport': 0,
        'OSTALO (A, N, O)_N1': 21.3860851946291,
        'OSTALO (A, N, O)_Nova S': 15.6304908675894,
        'OSTALO (A, N, O)_CAS Others (w/o SK, Nova Sport, Brainz)': 100 - 21.3860851946291 - 15.6304908675894,
        'OSTALO (A, N, O)_Grand Kanal': 10.3894919441524,
        'OSTALO (A, N, O)_Cinemania': 5.59380086843299,
        'OSTALO (A, N, O)_IDJ': 4.1480328297402,
        'OSTALO (A, N, O)_Pikaboo': 8.94222055926149,
        'OSTALO (A, N, O)_Vavoom': 6.44536297041165,
        'OSTALO (A, N, O)_CineStar': 2.99396406501471,
        'OSTALO (A, N, O)_CineStar Action': 4.43002716880657,
        'OSTALO (A, N, O)_Discovery': 2.27895590936159,
        'OSTALO (A, N, O)_Discovery ID': 0.658549153991927,
        'OSTALO (A, N, O)_HGTV': 0,
        'OSTALO (A, N, O)_TLC': 2.15294209026926,
        'OSTALO (A, N, O)_DIVA - Universal': 3.9686669244627,
        'OSTALO (A, N, O)_AXN': 1.61457711289044,
        'OSTALO (A, N, O)_Nickelodeon': 6.38516230710189,
        'OSTALO (A, N, O)_TV 1000': 1.04293807299414,
        'OSTALO (A, N, O)_Animal Planet': 1.87614264750958,
        'OSTALO (A, N, O)_NickJr': 0.0625893133800317,
        'OSTALO (A, N, O)_SK 1': 0,
        'OSTALO (A, N, O)_SK 2': 0,
        'OSTALO (A, N, O)_SK 3': 0,
        'OSTALO (A, N, O)_Nova Sport': 0
    }

    cas_others_dm_pool = ['CAS Others (w/o SK, Nova Sport, Brainz)','Grand Kanal','Cinemania','IDJ','Pikaboo','Vavoom','CineStar','CineStar Action','Discovery','Discovery ID','HGTV','TLC','DIVA - Universal','AXN','Nickelodeon','TV 1000','Animal Planet','NickJr']

    cas_non_agb = ['CAS "non AGB" (SK, Nova Sport, Brainz)','Brainz','Nova Sport','Sport klub SRB']

    cas_others = ['CAS Others (w/o SK, Nova Sport, Brainz)','Grand Kanal','Cinemania','IDJ','SK 1','SK 2','SK 3','Pikaboo','Vavoom','CineStar','CineStar Action','Discovery','Discovery ID','HGTV','TLC','DIVA - Universal','AXN','Nickelodeon','TV 1000','NickJr','Nova Sport','Animal Planet']

    cas_dm_pool_channels = ['N1','Nova S'] + cas_others_dm_pool

    cas_other_channels = ['N1','Nova S'] + cas_others

    dm_pool_shr_channels = cas_dm_pool_channels + cas_non_agb

    def get_collection(self,sheet,columns):
        topics = []
        for i, cellObj in enumerate(sheet, 2):
                tup = ()
                for col in columns:
                    val = sheet.cell(row=i,column=col).value
                    if val != None:
                        if type(val) is datetime.time:
                            val = val.isoformat()
                        tup = tup + (val,)
                if len(tup) != 0:
                    topics.append(tup)
        return topics

    # Return List type []
    def column_values(self,sheet,column):
        topics = []
        for i, cellObj in enumerate(sheet, 2):
            val = sheet.cell(row=i,column=column).value
            if val != None:
                topics.append(val)
        return sorted(list(set(topics)))


    def get_collection_start_from_3(self,sheet,columns):
        topics = []
        for i, cellObj in enumerate(sheet, 3):
                tup = ()
                for col in columns:
                    val = sheet.cell(row=i,column=col).value
                    if val != None:
                        tup = tup + (val,)
                if len(tup) != 0:
                    topics.append(tup)
        return topics

    def customers(self):
        return self.column_values(self.grp_baza,20)

    def channels(self):
        return self.column_values(self.e2_baza,30)

    def e2_collection(self):
        return self.get_collection(self.e2_baza,[28,3,20,30])

    def cas_collection(self):
        return self.get_collection_start_from_3(self.cas_baza,[3,4,5,2])

    def e2_channels(self):
        e2_channels = self.channels().copy()
        e2_channels.remove('Cas Media')
        return e2_channels

    def cas_channels(self):
        return ['Cas Media', 'CAS (SK, Nova Sport, Brainz)']

    def sk_nova(self):
        return ('Sport klub SRB', 'Nova Sport', 'Brainz')

    def collection(self,db,range):
        return self.get_collection(db,range)


