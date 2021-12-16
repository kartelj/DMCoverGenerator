# -*- coding: utf-8 -*- 
import tables.cover as cov
import tables.cover_cas as cov_cas
# import create_images_from_covers as cifc
import excel2img
from pathlib import Path

def add_rows(data, size):
    for i in range(size):
        data.append([])

# def export_images(cover, cover_cas):
#     print("Export images")
#     config = cov.f.cfg.Config()
#     cifc.exportCoverImages(config.NEW_FILE_PATH,cover,cover_cas)
                        

def exportCoverImages(csvName,cover,cover_cas):
    csvDirMarko = csvName.replace('.','_')+'/cover'
    Path(csvDirMarko).mkdir(parents=True, exist_ok=True)

    ###############
    #### Cover ####
    ###############

    coverMarkoImages = list(cover.keys())
    coverMarkoRanges = list(cover.values())
    coverMarkoImagesWithDir = [csvDirMarko+'/'+x for x in coverMarkoImages]

    excel2img.export_imgs(csvName, coverMarkoImagesWithDir ,"COVER Marko", coverMarkoRanges)

    ###################
    #### Cover Cas ####
    ###################

    csvDirMarkoCas = csvName.replace('.','_')+'/cover_cas'
    Path(csvDirMarkoCas).mkdir(parents=True, exist_ok=True)

    coverMarkoCasImages = list(cover_cas.keys())
    coverMarkoCasRanges = list(cover_cas.values())
    coverMarkoCasImagesWithDir = [csvDirMarkoCas+'/'+x for x in coverMarkoCasImages]

    excel2img.export_imgs(csvName, coverMarkoCasImagesWithDir ,"COVER Cas Marko", coverMarkoCasRanges)

def run():
    config = cov.f.cfg.Config()
    wb = config.wb
    ws = config.ws
    ws2 = config.ws2
    ws3 = config.ws3
    ws4 = config.ws4

    print('START')
    print('Ucitava covers')
    # Total
    coll = config.collection(config.grp_baza,[8,4,3,20,17])
    cover = cov.Cover(['grp'],coll)

    # 18-50
    coll2 = config.collection(config.grp_baza,[12,4,3,20,17])
    cover2 = cov.Cover(['grp'],coll2)

    # Duration
    coll3 = config.collection(config.grp_baza,[24,4,3,20,17])
    cover3 = cov.Cover(['duration'],coll3)

    # SOI
    cover4 = cov.Cover(['soi'],coll)

    cover5 = cov.Cover(['grp','soi'],coll)

    cover6 = cov.Cover(['grp','soi'],coll2)

    sov1 = cover.sov1()
    sov2 = cover2.sov2()
    print('Covers ucitani')


    print('kreira total tabelu')
    data = []
    cover_rows = {}

    data += cover.total_table()
    cover_rows['total'] = [1,len(data)]

    add_rows(data,3)

    print('kreira sov tabelu')

    start = len(data)+1
    data += cover.sov_table()
    cover_rows['sov_total'] = [start,len(data)]

    add_rows(data,3)

    print('kreira shr sov total tabelu')
    start = len(data)+1
    data += cover.shr_sov_total_table()
    cover_rows['shr_sov'] = [start,len(data)]

    add_rows(data,3)

    print('kreira 18-50 total tabelu')
    start = len(data)+1
    data += cover2.table_18_50()
    cover_rows['18-50'] = [start,len(data)]

    add_rows(data,3)

    print('kreira sov 18-50 tabelu')
    start = len(data)+1
    data += cover2.sov_18_50()
    cover_rows['sov_18-50'] = [start,len(data)]

    add_rows(data,3)

    print('kreira shr 18-50 tabelu')
    start = len(data)+1
    data += cover.shr_18_50_total_table()
    cover_rows['shr_18-50'] = [start,len(data)]

    add_rows(data,3)
    print('kreira Kontrolni duration tabelu')
    start = len(data)+1
    data += cover3.kontrolni_duration()
    cover_rows['kontrolni_duration'] = [start,len(data)]

    add_rows(data,3)
    print('kreira DURATION tabelu')
    duration_res = cover3.duration()
    start = len(data)+1
    data += duration_res['data']
    cover_rows['duration'] = [start,len(data)]
    duration_header = duration_res['cas_header']

    add_rows(data,3)
    print('kreira SOI total tabelu')
    soi_result = cover4.soi_total_table()
    start = len(data)+1
    data += soi_result['data']
    cover_rows['soi'] = [start,len(data)]
    soi_header = soi_result['dmp_header']
    new_dmp_header = soi_result['new_dmp_header']

    add_rows(data,3)
    print('kreira cpp 4+ Total tabelu')
    start = len(data)+1
    data += cover5.cpp_4_plus_table()
    cover_rows['cpp_4+'] = [start,len(data)]

    add_rows(data,3)
    print('kreira cpp 18-50 total tabelu')
    start = len(data)+1
    data += cover6.cpp_18_50_table()
    cover_rows['cpp_18_50'] = [start,len(data)]

    add_rows(data,3)
    print('kreira RATIO total tabelu')
    start = len(data)+1
    data += cover4.ratio_total_table()
    cover_rows['ratio'] = [start,len(data)]

    add_rows(data,3)
    print('kreira RATIO 18-50 tabelu')
    start = len(data)+1
    data += cover4.ratio_18_50_table()
    cover_rows['ratio_18_50'] = [start,len(data)]



    print('kreira MARKO SOV tabelu')
    data2 = []
    cover_marko_rows = {}

    data2 += cover.marko_sov1_sov2_table(['DM Pool', 'OSTALO (A, N, O)'],sov1,sov2)
    cover_marko_rows['sov1'] = [1,len(data2)]
    add_rows(data2,3)

    start = len(data2)+1
    data2 += cover.marko_sov1_sov2_table(['MEDIA House', 'MEDIA Pool'],sov1,sov2)
    cover_marko_rows['sov2'] = [start,len(data2)]

    add_rows(data2,3)

    start = len(data2)+1
    data2 += cover.marko_sov1_sov2_table(['MEDIACOM', 'ONE MEDIA'],sov1,sov2)
    cover_marko_rows['sov3'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO SHR duration tabelu')
    start = len(data2)+1
    data2 += cover3.marko_shr_duration(['DM Pool', 'MEDIA House', 'MEDIA Pool'])
    cover_marko_rows['shr1'] = [start,len(data2)]

    add_rows(data2,3)

    start = len(data2)+1
    data2 += cover3.marko_shr_duration(['MEDIACOM', 'ONE MEDIA', 'OSTALO (A, N, O)'])
    cover_marko_rows['shr2'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO Sh Total tabelu')
    start = len(data2)+1
    data2 += cover.marko_sh_total_population()
    cover_marko_rows['sh'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO Sh 18-50 tabelu')
    start = len(data2)+1
    data2 += cover.marko_sh_18_50_table()
    cover_marko_rows['sh_18_50'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO SOI tabelu')
    start = len(data2)+1
    data2 += cover4.marko_soi_table()
    cover_marko_rows['soi'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO CPP 4+ tabelu')
    start = len(data2)+1
    data2 += cover5.marko_cpp_4_plus()
    cover_marko_rows['cpp_4+'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO CPP 18-50 tabelu')
    start = len(data2)+1
    data2 += cover6.marko_cpp_18_50()
    cover_marko_rows['cpp_18_50'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO RATIO 4+ tabelu')
    start = len(data2)+1
    data2 += cover4.marko_ratio_4_plus()
    cover_marko_rows['ratio_4'] = [start,len(data2)]

    add_rows(data2,3)
    print('kreira MARKO RATIO 18-50 tabelu')
    start = len(data2)+1
    data2 += cover4.marko_ratio_18_50()
    cover_marko_rows['ratio_18_50'] = [start,len(data2)]

    ###################
    #### COVER CAS ####
    ###################

    cas_coll = config.collection(config.grp_baza,[8,4,3,20,2])
    cas_cover = cov_cas.CoverCas(['grp'],cas_coll)


    cas_coll2 = config.collection(config.grp_baza,[12,4,3,20,2])
    cas_cover2 = cov_cas.CoverCas(['grp'],cas_coll2)

    cas_coll3 = config.collection(config.grp_baza,[24,4,3,20,2])
    cas_cover3 = cov_cas.CoverCas(['duration'],cas_coll3)


    cas_cover4 = cov_cas.CoverCas(['soi'],cas_coll)

    cas_cover5 = cov_cas.CoverCas(['grp','soi'],cas_coll)

    cas_cover6 = cov_cas.CoverCas(['grp','soi'],cas_coll2)

    data3 = []
    cover_cas_rows = {}
    data3 += cas_cover.total_table()
    cover_cas_rows['total'] = [1,len(data3)]

    add_rows(data3,3)

    start = len(data3)+1
    data3 += cas_cover.sov_table()
    cover_cas_rows['sov'] = [start,len(data3)]
    add_rows(data3,3)

    start = len(data3)+1
    data3 += cas_cover.shr_sov_total_table()
    cover_cas_rows['shr'] = [start,len(data3)]

    add_rows(data3,3)

    start = len(data3)+1
    data3 += cas_cover2.table_18_50()
    cover_cas_rows['18_50'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover2.sov_18_50()
    cover_cas_rows['sov_18_50'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover.shr_18_50_total_table()
    cover_cas_rows['shr_18_50'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover3.kontrolni_duration()
    cover_cas_rows['kontrolni_duration'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover3.duration(duration_header)
    cover_cas_rows['duration'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover4.soi_total_table(new_dmp_header)
    cover_cas_rows['soi'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover5.cpp_4_plus_table()
    cover_cas_rows['cpp_4+'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover6.cpp_18_50_table()
    cover_cas_rows['cpp_18_50'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover4.ratio_total_table(new_dmp_header)
    cover_cas_rows['ratio'] = [start,len(data3)]

    add_rows(data3,3)
    start = len(data3)+1
    data3 += cas_cover4.ratio_18_50_table(new_dmp_header)
    cover_cas_rows['ratio_18_50'] = [start,len(data3)]


    print('kreira MARKO SOV tabelu')
    data4 = []
    cover_cas_marko_rows = {}
    sov1 = cas_cover.sov()
    sov2 = cas_cover2.sov()

    data4 += cas_cover.marko_sov1_sov2_table(['DM Pool','MEDIA House','MEDIA Pool'],sov1,sov2)
    cover_cas_marko_rows['sov1'] = [1,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover.marko_sov1_sov2_table(['MEDIACOM','ONE MEDIA','OSTALO (A, N, O)'],sov1,sov2)
    cover_cas_marko_rows['sov2'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover3.marko_duration(duration_header,['DM Pool','MEDIA House','MEDIA Pool'])
    cover_cas_marko_rows['duration_1'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover3.marko_duration(duration_header,['MEDIACOM','ONE MEDIA','OSTALO (A, N, O)'])
    cover_cas_marko_rows['duration_2'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover.marko_shr_table()
    cover_cas_marko_rows['shr'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover4.marko_soi_dm(new_dmp_header)
    cover_cas_marko_rows['soi'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    t1 = cas_cover5.marko_average_cpp()
    t2 = cas_cover6.marko_average_cpp()
    data4 += cas_cover.merge_tables(t1,t2)
    cover_cas_marko_rows['cpp'] = [start,len(data4)]

    add_rows(data4,3)
    start = len(data4)+1
    data4 += cas_cover4.marko_ratio_total(new_dmp_header)
    cover_cas_marko_rows['ratio'] = [start,len(data4)]
    

    for row in data:
        ws.append(row)

    for row in data2:
        ws2.append(row)

    for row in data3:
        ws3.append(row)

    for row in data4:
        ws4.append(row)


    cov.style.cover_style(ws,cover_rows)
    cover_img = cov.style.marko_cover_style(ws2,cov.f.Functions().QUARTER,cover_marko_rows)
    cov.style.cover_cas_style(ws3,cover_cas_rows)
    cover_cas_img = cov.style.marko_cover_cas_style(ws4,cov.f.Functions().QUARTER,cover_cas_marko_rows)

    print('SAVING EXCEL FILE!')
    wb.save(config.NEW_FILE_PATH)

    exportCoverImages(config.NEW_FILE_PATH, cover_img,cover_cas_img)

    print('KRAJ!')
if __name__ == "__main__":
    run()
