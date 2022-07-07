# -*- codeing = utf-8 -*-
# @Time: 2021/3/11 19:53
# @Author: Foxhuty
# @File: score_data_analysis.py
# @Software: PyCharm
import pandas as pd
import numpy as np
import time
import os

pd.options.mode.use_inf_as_na = True
pd.set_option('display.unicode.east_asian_width', True)
pd.set_option('display.max_rows', 1000)


# pd.set_option('display.max_columns',1000)


class ContrastScores(object):
    path_one = None
    path_two = None

    def __init__(self, path_this=None, path_last=None):
        self.path_this = path_this
        self.path_last = path_last

    def get_df_contrast(self, exam1=None, exam2=None):
        """
        计算取得文科与理科在两次考试中的成绩对比，输出为excel表格
        :param exam1: 本次考试名称，如期末，半期，一诊，二诊等
        :param exam2: 作为对比的上一次考试名称，如期末，半期，一诊，二诊等
        :return: 生成excel电子表格，无返回值
        """
        df_arts_this = pd.read_excel(self.path_this, sheet_name='文科', dtype={'班级': str, '考生号': str, '考号': str})
        df_arts_this['名次'] = df_arts_this['总分'].rank(method='min', ascending=False)
        df_arts_last = pd.read_excel(self.path_last, sheet_name='文科', dtype={'班级': str, '考生号': str, '考号': str})
        df_arts_last['名次'] = df_arts_last['总分'].rank(method='min', ascending=False)
        df_arts_last = df_arts_last.loc[:, ['姓名', '总分', '名次']]
        df_contrast = df_arts_this.merge(df_arts_last, on='姓名', how='left')
        df_contrast['变化'] = df_contrast['名次_y'] - df_contrast['名次_x']
        df_contrast.rename(columns={'总分_x': exam1 + '总分', '名次_x': exam1 + '名次',
                                    '总分_y': exam2 + '总分', '名次_y': exam2 + '名次'}, inplace=True)

        class_names = df_contrast['班级'].unique()

        df_science_this = pd.read_excel(self.path_this, sheet_name='理科', dtype={'班级': str, '考生号': str, '考号': str})
        df_science_this['名次'] = df_science_this['总分'].rank(method='min', ascending=False)
        df_science_last = pd.read_excel(self.path_last, sheet_name='理科', dtype={'班级': str, '考生号': str, '考号': str})
        df_science_last['名次'] = df_science_last['总分'].rank(method='min', ascending=False)
        df_science_last = df_science_last.loc[:, ['姓名', '总分', '名次']]
        df_science_contrast = df_science_this.merge(df_science_last, on='姓名', how='left')
        df_science_contrast['变化'] = df_science_contrast['名次_y'] - df_science_contrast['名次_x']
        df_science_contrast.rename(columns={'总分_x': exam1 + '总分', '名次_x': exam1 + '名次',
                                            '总分_y': exam2 + '总分', '名次_y': exam2 + '名次'}, inplace=True)
        class_names_science = df_science_contrast['班级'].unique()
        writer = pd.ExcelWriter(r'D:\成绩统计结果\与上次考试成绩对比表.xlsx')
        for i in class_names:
            class_name = df_contrast[df_contrast['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name.to_excel(writer, sheet_name=i, index=False)
        for i in class_names_science:
            class_name_science = df_science_contrast[df_science_contrast['班级'] == i].reset_index(drop=True)
            class_name_science['序号'] = [k + 1 for k in class_name_science.index]
            class_name_science.to_excel(writer, sheet_name=i, index=False)

        df_contrast.to_excel(writer, sheet_name='文科对比表', index=False)
        df_science_contrast.to_excel(writer, sheet_name='理科对比表', index=False)
        writer.close()

    @classmethod
    def contrast(cls, exam1=None, exam2=None):
        """
        计算取得文科与理科在两次考试中的成绩对比，输出为excel表格
        :param exam1: 本次考试名称，如期末，半期，一诊，二诊等
        :param exam2: 作为对比的上一次考试名称，如期末，半期，一诊，二诊等
        :return: 生成excel电子表格，无返回值
        """
        df_arts_this = pd.read_excel(cls.path_one, sheet_name='文科', dtype={'班级': str, '考生号': str, '考号': str})
        df_arts_this = df_arts_this.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '政治', '历史', '地理', '总分']]
        df_arts_this['名次'] = df_arts_this['总分'].rank(method='min', ascending=False)
        df_arts_last = pd.read_excel(cls.path_two, sheet_name='文科', dtype={'班级': str, '考生号': str, '考号': str})
        df_arts_last['名次'] = df_arts_last['总分'].rank(method='min', ascending=False)
        df_arts_last = df_arts_last.loc[:, ['姓名', '总分', '名次']]
        df_contrast = df_arts_this.merge(df_arts_last, on='姓名', how='left')
        df_contrast['变化'] = df_contrast['名次_y'] - df_contrast['名次_x']
        df_contrast.rename(columns={'总分_x': exam1 + '总分', '名次_x': exam1 + '名次',
                                    '总分_y': exam2 + '总分', '名次_y': exam2 + '名次'}, inplace=True)
        class_names = df_contrast['班级'].unique()

        df_science_this = pd.read_excel(cls.path_one, sheet_name='理科', dtype={'班级': str, '考生号': str, '考号': str})
        df_science_this = df_science_this.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '物理', '化学', '生物', '总分']]
        df_science_this['名次'] = df_science_this['总分'].rank(method='min', ascending=False)
        df_science_last = pd.read_excel(cls.path_two, sheet_name='理科', dtype={'班级': str, '考生号': str, '考号': str})
        df_science_last['名次'] = df_science_last['总分'].rank(method='min', ascending=False)
        df_science_last = df_science_last.loc[:, ['姓名', '总分', '名次']]
        df_science_contrast = df_science_this.merge(df_science_last, on='姓名', how='left')
        df_science_contrast['变化'] = df_science_contrast['名次_y'] - df_science_contrast['名次_x']
        df_science_contrast.rename(columns={'总分_x': exam1 + '总分', '名次_x': exam1 + '名次',
                                            '总分_y': exam2 + '总分', '名次_y': exam2 + '名次'}, inplace=True)
        class_names_science = df_science_contrast['班级'].unique()

        writer = pd.ExcelWriter(r'D:\成绩统计结果\与上次考试成绩对比表.xlsx')
        for i in class_names:
            class_name = df_contrast[df_contrast['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name.to_excel(writer, sheet_name=i, index=False)
        for i in class_names_science:
            class_name_science = df_science_contrast[df_science_contrast['班级'] == i].reset_index(drop=True)
            class_name_science['序号'] = [k + 1 for k in class_name_science.index]
            class_name_science.to_excel(writer, sheet_name=i, index=False)
        df_contrast.to_excel(writer, sheet_name='文科对比表', index=False)
        df_science_contrast.to_excel(writer, sheet_name='理科对比表', index=False)
        writer.close()


class ScoreAnalysis(object):
    """
    高中各类考试成绩分析，用于计算平均分，有效分，有效分人数，错位生人数，成绩对比，学科评定，考室安排，学生个人成绩单等。
    简单高效地完成考试成绩分析。--
    """
    arts_scores = []
    science_scores = []
    top_n = None
    numbers_in_room = None

    # order_arts = ['语文', '数学', '英语', '政治', '历史', '地理', '总分']
    # order_science = ['语文', '数学', '英语', '物理', '化学', '生物', '总分']

    def __init__(self, path):
        self.path = path
        self.df_arts = pd.read_excel(self.path, sheet_name='文科', index_col='序号',
                                     dtype={
                                         '班级': str,
                                         '序号': str,
                                         '名次': str,
                                         '考生号': str,
                                         '考号': str

                                     }
                                     )
        self.df_science = pd.read_excel(self.path, sheet_name='理科', index_col='序号',
                                        dtype={
                                            '班级': str,
                                            '序号': str,
                                            '名次': str,
                                            '考生号': str,
                                            '考号': str

                                        }
                                        )

    def __str__(self):
        return f'正在对{os.path.basename(self.path)}进行成绩分析处理'

    def get_av(self):
        """
        计算各科平均分
        :return: 文科，理科各班各科平均分
        """
        av_arts = self.get_av_arts()
        av_science = self.get_av_science()
        av_arts_percentage, av_science_percentage = self.get_av_diagram()
        av_percentage_arts = pd.concat([av_arts, av_arts_percentage], axis=1)
        av_percentage_science = pd.concat([av_science, av_science_percentage], axis=1)
        arts_order = ['参考人数', '语文', '语文占比', '数学', '数学占比', '英语', '英语占比',
                      '政治', '政治占比', '历史', '历史占比', '地理', '地理占比', '总分', '总分占比']
        science_order = ['参考人数', '语文', '语文占比', '数学', '数学占比', '英语', '英语占比',
                         '物理', '物理占比', '化学', '化学占比', '生物', '生物占比', '总分', '总分占比']
        av_percentage_arts = av_percentage_arts[arts_order]
        av_percentage_science = av_percentage_science[science_order]
        return av_percentage_arts, av_percentage_science

    def get_av_science(self):
        av_class_science = self.df_science.groupby(['班级'])[['语文', '数学', '英语', '物理', '化学', '生物', '总分']].mean()
        av_general_science = self.df_science[['语文', '数学', '英语', '物理', '化学', '生物', '总分']].apply(np.nanmean, axis=0)
        av_general_science.name = '年级'
        av_science = av_class_science.append(av_general_science)
        science_students = self.get_student_number_class(self.df_science)
        av_science = av_science.join(science_students)
        order_science = ['参考人数', '语文', '数学', '英语', '物理', '化学', '生物', '总分']
        av_science = av_science[order_science]
        return av_science

    def get_av_arts(self):
        av_class_arts = self.df_arts.groupby(['班级'])[['语文', '数学', '英语', '政治', '历史', '地理', '总分']].mean()
        av_general_arts = self.df_arts[['语文', '数学', '英语', '政治', '历史', '地理', '总分']].apply(np.nanmean, axis=0)
        # av_general_arts.name = '年级'
        # av_arts = av_class_arts.append(av_general_arts)
        # 新版pandas取消append方法后，代码如下：
        av_class_arts.loc['年级'] = av_general_arts

        arts_students = self.get_student_number_class(self.df_arts)
        # av_arts = av_arts.join(arts_students)
        av_arts = av_class_arts.join(arts_students)
        order = ['参考人数', '语文', '数学', '英语', '政治', '历史', '地理', '总分']
        av_arts = av_arts[order]
        return av_arts

    def get_av_subjects(self, df_data, *args):
        av = df_data.groupby(['班级'])[[*args]].mean()
        av_general = df_data[[*args]].apply(np.nanmean, axis=0)
        av_general.name = '年级'
        av = av.append(av_general)
        student_num = self.get_student_number_class(df_data)
        av = av.join(student_num)
        return av

    def get_av_diagram(self):
        av_class_arts = self.df_arts.groupby(['班级'])[['语文', '数学', '英语', '政治', '历史', '地理', '总分']].mean()
        av_general_arts = self.df_arts[['语文', '数学', '英语', '政治', '历史', '地理', '总分']].apply(np.nanmean, axis=0)
        av_diagram_arts = av_class_arts / av_general_arts
        av_class_science = self.df_science.groupby(['班级'])[['语文', '数学', '英语', '物理', '化学', '生物', '总分']].mean()
        av_general_science = self.df_science[['语文', '数学', '英语', '物理', '化学', '生物', '总分']].apply(np.nanmean, axis=0)
        av_diagram_science = av_class_science / av_general_science
        # 新增一列，用百分数表示班级在年级平均分中的占比
        av_diagram_arts = av_diagram_arts.applymap(lambda x: format(x, '.2%'))
        av_diagram_science = av_diagram_science.applymap(lambda x: format(x, '.2%'))
        av_diagram_arts.rename(columns={'语文': '语文占比', '数学': '数学占比', '英语': '英语占比',
                                        '政治': '政治占比', '历史': '历史占比', '地理': '地理占比',
                                        '总分': '总分占比'}, inplace=True)
        av_diagram_science.rename(columns={'语文': '语文占比', '数学': '数学占比', '英语': '英语占比',
                                           '物理': '物理占比', '化学': '化学占比', '生物': '生物占比',
                                           '总分': '总分占比'}, inplace=True)

        with pd.ExcelWriter(r'D:\成绩统计结果\平均分占比率.xlsx') as writer:
            av_diagram_arts.to_excel(writer, sheet_name='文科')
            av_diagram_science.to_excel(writer, sheet_name='理科')









        return av_diagram_arts, av_diagram_science

    @staticmethod
    def av_subject_percentage(data, subject):
        data[subject + '占比'] = data[subject].apply(lambda x: format(x, '.2%'))
        return data

    def exam_room_info(self):
        """
        计算生成文理科各考室学生名单
        :return:
        """
        self.df_arts.sort_values(by='总分', ascending=False, inplace=True)
        self.df_science.sort_values(by='总分', ascending=False, inplace=True)
        if len(self.df_arts) % ScoreAnalysis.numbers_in_room != 0:
            room_numbers = [f'文科第{str(i + 1)}考室' for i in
                            list(range(len(self.df_arts) // ScoreAnalysis.numbers_in_room + 1))]
        else:
            room_numbers = [f'文科第{str(i + 1)}考室' for i in
                            list(range(len(self.df_arts) // ScoreAnalysis.numbers_in_room))]
        if len(self.df_science) % ScoreAnalysis.numbers_in_room != 0:
            room_numbers_science = [f'理科第{str(i + 1)}考室' for i in
                                    list(range(len(self.df_science) // ScoreAnalysis.numbers_in_room + 1))]
        else:
            room_numbers_science = [f'理科第{str(i + 1)}考室' for i in
                                    list(range(len(self.df_science) // ScoreAnalysis.numbers_in_room))]
        print(room_numbers)
        print(room_numbers_science)
        df_arts = self.df_arts.copy()
        df_science = self.df_science.copy()
        df_arts['考室号'] = None
        df_arts['座位号'] = None
        df_arts = df_arts.loc[:, ['班级', '姓名', '考号', '考室号', '座位号']]
        df_arts.reset_index(drop=True, inplace=True)

        df_science['考室号'] = None
        df_science['座位号'] = None
        df_science = df_science.loc[:, ['班级', '姓名', '考号', '考室号', '座位号']]
        df_science.reset_index(drop=True, inplace=True)

        df_room_students = []
        df_room_students_science = []
        arts = pd.DataFrame(columns=df_arts.columns)
        science = pd.DataFrame(columns=df_science.columns)
        for idx, room_number in enumerate(room_numbers):
            begin = idx * ScoreAnalysis.numbers_in_room
            end = begin + ScoreAnalysis.numbers_in_room
            df_room_student = df_arts.iloc[begin:end]
            df_room_students.append((idx, room_number, df_room_student))
            # print(df_room_students)
        writer = pd.ExcelWriter(r'D:\成绩统计结果\文理科考室学生名单.xlsx')
        for idx, room_number, df_room_student in df_room_students:
            for i in df_room_student.index:
                df_room_student = df_room_student.copy()
                df_room_student['考室号'].at[i] = room_number
                df_room_student['座位号'].at[
                    i] = i + 1 if i < ScoreAnalysis.numbers_in_room else i - idx * ScoreAnalysis.numbers_in_room + 1
            df_room_student.to_excel(writer, sheet_name=room_number, index=False)
            arts = arts.append(df_room_student)
        for idx, room_number in enumerate(room_numbers_science):
            begin = idx * ScoreAnalysis.numbers_in_room
            end = begin + ScoreAnalysis.numbers_in_room
            df_room_student_science = df_science.iloc[begin:end]
            df_room_students_science.append((idx, room_number, df_room_student_science))
        for idx, room_number, df_room_student_science in df_room_students_science:
            for i in df_room_student_science.index:
                df_room_student_science = df_room_student_science.copy()
                df_room_student_science['考室号'].at[i] = room_number
                df_room_student_science['座位号'].at[
                    i] = i + 1 if i < ScoreAnalysis.numbers_in_room else i - idx * ScoreAnalysis.numbers_in_room + 1
            df_room_student_science.to_excel(writer, sheet_name=room_number, index=False)
            science = science.append(df_room_student_science)
        df_arts_general = arts.sort_values(by=['班级', '考室号'], ascending=[True, True])
        df_science_general = science.sort_values(by=['班级', '考室号'], ascending=[True, True])
        df_arts_general.to_excel(writer, sheet_name='文科', index=False)
        df_science_general.to_excel(writer, sheet_name='理科', index=False)
        # 计算考生座签
        df_seat_arts = arts.copy()
        df_seat_science = science.copy()
        for i in df_seat_arts.index:
            df_seat_arts.loc[i + 0.5] = ['班级', '姓名', '考号', '考室号', '座位号']
            # df2.to_excel(f'{df2.iloc[0,0]}.xlsx',index=False)
        df_seat_arts.sort_index(inplace=True, ignore_index=True)
        for i in df_seat_science.index:
            df_seat_science.loc[i + 0.5] = ['班级', '姓名', '考号', '考室号', '座位号']
        df_seat_science.sort_index(inplace=True, ignore_index=True)
        df_seat_arts.to_excel(writer, sheet_name='文科座签', index=False)
        df_seat_science.to_excel(writer, sheet_name='理科座签', index=False)

        writer.close()
        print('successfully done')

    def top_n_students(self):

        top_chn = self.get_top_n(self.df_arts, '语文', n=ScoreAnalysis.top_n)
        top_math = self.get_top_n(self.df_arts, '数学', n=ScoreAnalysis.top_n)
        top_eng = self.get_top_n(self.df_arts, '英语', n=ScoreAnalysis.top_n)
        top_pol = self.get_top_n(self.df_arts, '政治', n=ScoreAnalysis.top_n)
        top_his = self.get_top_n(self.df_arts, '历史', n=ScoreAnalysis.top_n)
        top_geo = self.get_top_n(self.df_arts, '地理', n=ScoreAnalysis.top_n)
        top_total = self.get_top_n(self.df_arts, '总分', n=ScoreAnalysis.top_n)
        top_arts = pd.concat([top_chn, top_math, top_eng, top_pol, top_his, top_geo, top_total], axis=1)
        top_arts.index = [i + 1 for i in top_arts.index]

        top_chn_science = self.get_top_n(self.df_science, '语文', n=ScoreAnalysis.top_n)
        top_math_science = self.get_top_n(self.df_science, '数学', n=ScoreAnalysis.top_n)
        top_eng_science = self.get_top_n(self.df_science, '英语', n=ScoreAnalysis.top_n)
        top_phy_science = self.get_top_n(self.df_science, '物理', n=ScoreAnalysis.top_n)
        top_chem_science = self.get_top_n(self.df_science, '化学', n=ScoreAnalysis.top_n)
        top_bio_science = self.get_top_n(self.df_science, '生物', n=ScoreAnalysis.top_n)
        top_total_science = self.get_top_n(self.df_science, '总分', n=ScoreAnalysis.top_n)

        top_science = pd.concat([top_chn_science, top_math_science, top_eng_science,
                                 top_phy_science, top_chem_science, top_bio_science, top_total_science], axis=1)
        top_science.index = [i + 1 for i in top_science.index]
        with pd.ExcelWriter(r'D:\成绩统计结果\文理前N名.xlsx') as writer:
            top_arts.to_excel(writer, sheet_name='文科前N名', index_label='序号')
            top_science.to_excel(writer, sheet_name='理科前N名', index_label='序号')

    def get_goodscores_arts(self, goodtotal_arts):
        """
        计算文科各科有效分
        goodtotal:划线总分，高线，中线，低线
        """

        chn = self.get_subject_good_score(self.df_arts, '语文', goodtotal_arts)
        math = self.get_subject_good_score(self.df_arts, '数学', goodtotal_arts)
        eng = self.get_subject_good_score(self.df_arts, '英语', goodtotal_arts)
        pol = self.get_subject_good_score(self.df_arts, '政治', goodtotal_arts)
        his = self.get_subject_good_score(self.df_arts, '历史', goodtotal_arts)
        geo = self.get_subject_good_score(self.df_arts, '地理', goodtotal_arts)

        if (chn + math + eng + pol + his + geo) > goodtotal_arts:
            math -= 1
        if (chn + math + eng + pol + his + geo) < goodtotal_arts:
            eng += 1

        return chn, math, eng, pol, his, geo, goodtotal_arts

    def get_goodscores_science(self, goodtotal_science):
        """
        计算理科各科有效分
        goodtotal:划线总分，高线，中线，低线
        """
        chn = self.get_subject_good_score(self.df_science, '语文', goodtotal_science)
        math = self.get_subject_good_score(self.df_science, '数学', goodtotal_science)
        eng = self.get_subject_good_score(self.df_science, '英语', goodtotal_science)
        phys = self.get_subject_good_score(self.df_science, '物理', goodtotal_science)
        chem = self.get_subject_good_score(self.df_science, '化学', goodtotal_science)
        bio = self.get_subject_good_score(self.df_science, '生物', goodtotal_science)
        if (chn + math + eng + phys + chem + bio) > goodtotal_science:
            math -= 1
        if (chn + math + eng + phys + chem + bio) < goodtotal_science:
            eng += 1

        return chn, math, eng, phys, chem, bio, goodtotal_science

    def good_scores_arts_ratio(self, chn, math, eng, pol, his, geo, total):
        # 计算各班单有效学生人数
        single_chn_arts, double_chn_arts = self.get_single_double_score(self.df_arts, '语文', chn, total)
        single_math_arts, double_math_arts = self.get_single_double_score(self.df_arts, '数学', math, total)
        single_eng_arts, double_eng_arts = self.get_single_double_score(self.df_arts, '英语', eng, total)
        single_pol_arts, double_pol_arts = self.get_single_double_score(self.df_arts, '政治', pol, total)
        single_his_arts, double_his_arts = self.get_single_double_score(self.df_arts, '历史', his, total)
        single_geo_arts, double_geo_arts = self.get_single_double_score(self.df_arts, '地理', geo, total)
        single_total_arts, double_total_arts = self.get_single_double_score(self.df_arts, '总分', total, total)
        # 计算参考人数
        name_num = self.df_arts.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'
        result_single = pd.concat([name_num, single_chn_arts, single_math_arts, single_eng_arts,
                                   single_pol_arts, single_his_arts, single_geo_arts, single_total_arts],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_arts, double_math_arts, double_eng_arts,
             double_pol_arts, double_his_arts, double_geo_arts, double_total_arts], axis=1)

        result_single.loc['文科共计'] = [result_single['参考人数'].sum(),
                                     result_single['语文'].sum(),
                                     result_single['数学'].sum(),
                                     result_single['英语'].sum(),
                                     result_single['政治'].sum(),
                                     result_single['历史'].sum(),
                                     result_single['地理'].sum(),
                                     result_single['总分'].sum()
                                     ]
        # 新增上线率一列并用百分数表示
        result_single = self.good_scores_ratio(result_single, '总分')
        result_single = self.good_scores_ratio(result_single, '语文')
        result_single = self.good_scores_ratio(result_single, '数学')
        result_single = self.good_scores_ratio(result_single, '英语')
        result_single = self.good_scores_ratio(result_single, '政治')
        result_single = self.good_scores_ratio(result_single, '历史')
        result_single = self.good_scores_ratio(result_single, '地理')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '政治', '政治上线率', '历史', '历史上线率', '地理', '地理上线率', '总分', '总分上线率']
        result_single = result_single[order]
        result_double.loc['文科共计'] = [result_double['参考人数'].sum(),
                                     result_double['语文'].sum(),
                                     result_double['数学'].sum(),
                                     result_double['英语'].sum(),
                                     result_double['政治'].sum(),
                                     result_double['历史'].sum(),
                                     result_double['地理'].sum(),
                                     result_double['总分'].sum()]
        result_double = self.good_scores_ratio(result_double, '总分')
        result_double = self.good_scores_ratio(result_double, '语文')
        result_double = self.good_scores_ratio(result_double, '数学')
        result_double = self.good_scores_ratio(result_double, '英语')
        result_double = self.good_scores_ratio(result_double, '政治')
        result_double = self.good_scores_ratio(result_double, '历史')
        result_double = self.good_scores_ratio(result_double, '地理')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '政治', '政治上线率', '历史', '历史上线率', '地理', '地理上线率', '总分', '总分上线率']
        result_double = result_double[order]
        single_double_concat = pd.concat([result_single, result_double], keys=['单有效', '双有效'])

        return single_double_concat

    def goodscore_arts(self, chn, math, eng, pol, his, geo, total):
        """
        计算文科各科各班单有效和双有效人数
        """
        # 计算各班单,双有效学生人数
        single_chn_arts, double_chn_arts = self.get_single_double_score(self.df_arts, '语文', chn, total)
        single_math_arts, double_math_arts = self.get_single_double_score(self.df_arts, '数学', math, total)
        single_eng_arts, double_eng_arts = self.get_single_double_score(self.df_arts, '英语', eng, total)
        single_pol_arts, double_pol_arts = self.get_single_double_score(self.df_arts, '政治', pol, total)
        single_his_arts, double_his_arts = self.get_single_double_score(self.df_arts, '历史', his, total)
        single_geo_arts, double_geo_arts = self.get_single_double_score(self.df_arts, '地理', geo, total)
        single_total_arts, double_total_arts = self.get_single_double_score(self.df_arts, '总分', total, total)
        # 计算参考人数
        name_num = self.df_arts.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'
        #
        goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng,
                          '政治': pol, '历史': his, '地理': geo, '总分': total}
        goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

        result_single = pd.concat([name_num, single_chn_arts, single_math_arts, single_eng_arts,
                                   single_pol_arts, single_his_arts, single_geo_arts, single_total_arts],
                                  axis=1)

        result_double = pd.concat(
            [name_num, double_chn_arts, double_math_arts, double_eng_arts,
             double_pol_arts, double_his_arts, double_geo_arts, double_total_arts], axis=1)
        # 新增一行文科共计
        result_single.loc['文科共计'] = [result_single['参考人数'].sum(),
                                     result_single['语文'].sum(),
                                     result_single['数学'].sum(),
                                     result_single['英语'].sum(),
                                     result_single['政治'].sum(),
                                     result_single['历史'].sum(),
                                     result_single['地理'].sum(),
                                     result_single['总分'].sum()
                                     ]
        # 新增上线率一列并用百分数表示
        result_single = self.good_scores_ratio(result_single, '总分')

        # 新增一行文科共计。
        result_double.loc['文科共计'] = [result_double['参考人数'].sum(),
                                     result_double['语文'].sum(),
                                     result_double['数学'].sum(),
                                     result_double['英语'].sum(),
                                     result_double['政治'].sum(),
                                     result_double['历史'].sum(),
                                     result_double['地理'].sum(),
                                     result_double['总分'].sum()]
        # 计算错位生人数
        unmatched_dict = {'参考人数': name_num, '语文': single_total_arts - double_chn_arts,
                          '数学': single_total_arts - double_math_arts, '英语': single_total_arts - double_eng_arts,
                          '政治': single_total_arts - double_pol_arts, '历史': single_total_arts - double_his_arts,
                          '地理': single_total_arts - double_geo_arts, '总分': single_total_arts - double_total_arts}
        unmatched_df = pd.DataFrame(unmatched_dict)
        # 新增一行共计
        unmatched_df.loc['共计'] = [unmatched_df['参考人数'].sum(),
                                  unmatched_df['语文'].sum(),
                                  unmatched_df['数学'].sum(),
                                  unmatched_df['英语'].sum(),
                                  unmatched_df['政治'].sum(),
                                  unmatched_df['历史'].sum(),
                                  unmatched_df['地理'].sum(),
                                  unmatched_df['总分'].sum()
                                  ]
        # 合并所要数据：有效分数，单有效，双有效，错位数
        result_final_arts = pd.concat([goodscore_df, result_single, result_double, unmatched_df], axis=0,
                                      keys=['有效分数', '单有效', '双有效', '错位数'])
        result_final_arts.fillna(0, inplace=True)

        # 计算错位生名单
        df_chn = self.get_unmatched_students(self.df_arts, '语文', chn, total)
        df_math = self.get_unmatched_students(self.df_arts, '数学', math, total)
        df_eng = self.get_unmatched_students(self.df_arts, '英语', eng, total)
        df_pol = self.get_unmatched_students(self.df_arts, '政治', pol, total)
        df_his = self.get_unmatched_students(self.df_arts, '历史', his, total)
        df_geo = self.get_unmatched_students(self.df_arts, '地理', geo, total)
        unmatched_students_arts = pd.concat([df_chn, df_math, df_eng, df_pol, df_his, df_geo], axis=1)
        # 计算学科贡献率，命中率和等级评定
        shoot_dict = {'语文': result_double['语文'] / result_single['语文'],
                      '数学': result_double['数学'] / result_single['数学'],
                      '英语': result_double['英语'] / result_single['英语'],
                      '政治': result_double['政治'] / result_single['政治'],
                      '历史': result_double['历史'] / result_single['历史'],
                      '地理': result_double['地理'] / result_single['地理']}
        shoot_df = pd.DataFrame(shoot_dict)

        contribution_dict = {'语文': result_double['语文'] / result_double['总分'],
                             '数学': result_double['数学'] / result_double['总分'],
                             '英语': result_double['英语'] / result_double['总分'],
                             '政治': result_double['政治'] / result_double['总分'],
                             '历史': result_double['历史'] / result_double['总分'],
                             '地理': result_double['地理'] / result_double['总分']}
        contribution_df = pd.DataFrame(contribution_dict)

        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=shoot_df.columns, index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('政治')
        grade_assess('历史')
        grade_assess('地理')
        # 命中率和贡献率转化为百分数
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))
        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        final_grade = pd.concat([contribution_df, shoot_df, grade],
                                keys=['贡献率', '命中率', '等级'])

        return result_final_arts, final_grade, unmatched_students_arts

    def goodscore_science(self, chn, math, eng, pol, his, geo, total):
        """
        计算理科各科各班上单有效和双有效分人数
        """
        single_chn_science, double_chn_science = self.get_single_double_score(self.df_science, '语文', chn, total)
        single_math_science, double_math_science = self.get_single_double_score(self.df_science, '数学', math, total)
        single_eng_science, double_eng_science = self.get_single_double_score(self.df_science, '英语', eng, total)
        single_phys_science, double_phys_science = self.get_single_double_score(self.df_science, '物理', pol, total)
        single_chem_science, double_chem_science = self.get_single_double_score(self.df_science, '化学', his, total)
        single_bio_science, double_bio_science = self.get_single_double_score(self.df_science, '生物', geo, total)
        single_total_science, double_total_science = self.get_single_double_score(self.df_science, '总分', total, total)

        name_num = self.df_science.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        goodscore_dict = {'参考人数': ' ', '语文': chn, '数学': math, '英语': eng, '物理': pol,
                          '化学': his, '生物': geo, '总分': total}
        goodscore_df = pd.DataFrame(goodscore_dict, index=['有效分数'])

        result_single = pd.concat([name_num, single_chn_science, single_math_science, single_eng_science,
                                   single_phys_science, single_chem_science, single_bio_science, single_total_science],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_science, double_math_science, double_eng_science, double_phys_science,
             double_chem_science, double_bio_science, double_total_science], axis=1)

        result_single.loc['理科共计'] = [result_single['参考人数'].sum(),
                                     result_single['语文'].sum(),
                                     result_single['数学'].sum(),
                                     result_single['英语'].sum(),
                                     result_single['物理'].sum(),
                                     result_single['化学'].sum(),
                                     result_single['生物'].sum(),
                                     result_single['总分'].sum()
                                     ]
        # 新增一列上线率
        result_single = self.good_scores_ratio(result_single, '总分')

        result_double.loc['理科共计'] = [result_double['参考人数'].sum(),
                                     result_double['语文'].sum(),
                                     result_double['数学'].sum(),
                                     result_double['英语'].sum(),
                                     result_double['物理'].sum(),
                                     result_double['化学'].sum(),
                                     result_double['生物'].sum(),
                                     result_double['总分'].sum()]
        # 计算错位生人数
        unmatched_dict = {'参考人数': name_num, '语文': single_total_science - double_chn_science,
                          '数学': single_total_science - double_math_science,
                          '英语': single_total_science - double_eng_science,
                          '物理': single_total_science - double_phys_science,
                          '化学': single_total_science - double_chem_science,
                          '生物': single_total_science - double_bio_science,
                          '总分': single_total_science - double_total_science}
        unmatched_df = pd.DataFrame(unmatched_dict)
        unmatched_df.loc['共计'] = [unmatched_df['参考人数'].sum(),
                                  unmatched_df['语文'].sum(),
                                  unmatched_df['数学'].sum(),
                                  unmatched_df['英语'].sum(),
                                  unmatched_df['物理'].sum(),
                                  unmatched_df['化学'].sum(),
                                  unmatched_df['生物'].sum(),
                                  unmatched_df['总分'].sum()]
        result_final_science = pd.concat([goodscore_df, result_single, result_double, unmatched_df], axis=0,
                                         keys=['有效分数', '单有效', '双有效', '错位数'])
        result_final_science.fillna(0, inplace=True)
        # 计算学科错位生名单
        df_chn = self.get_unmatched_students(self.df_science, '语文', chn, total)
        df_math = self.get_unmatched_students(self.df_science, '数学', math, total)
        df_eng = self.get_unmatched_students(self.df_science, '英语', eng, total)
        df_pol = self.get_unmatched_students(self.df_science, '物理', pol, total)
        df_his = self.get_unmatched_students(self.df_science, '化学', his, total)
        df_geo = self.get_unmatched_students(self.df_science, '生物', geo, total)
        unmatched_students_science = pd.concat([df_chn, df_math, df_eng, df_pol, df_his, df_geo], axis=1)
        # 计算学科贡献率，命中率和等级评定
        shoot_dict = {'语文': result_double['语文'] / result_single['语文'],
                      '数学': result_double['数学'] / result_single['数学'],
                      '英语': result_double['英语'] / result_single['英语'],
                      '物理': result_double['物理'] / result_single['物理'],
                      '化学': result_double['化学'] / result_single['化学'],
                      '生物': result_double['生物'] / result_single['生物']}
        shoot_df = pd.DataFrame(shoot_dict)
        contribution_dict = {'语文': result_double['语文'] / result_single['总分'],
                             '数学': result_double['数学'] / result_single['总分'],
                             '英语': result_double['英语'] / result_single['总分'],
                             '物理': result_double['物理'] / result_single['总分'],
                             '化学': result_double['化学'] / result_single['总分'],
                             '生物': result_double['生物'] / result_single['总分']}
        contribution_df = pd.DataFrame(contribution_dict)
        result_single.fillna(0, inplace=True)
        result_double.fillna(0, inplace=True)
        shoot_df.fillna(0, inplace=True)
        contribution_df.fillna(0, inplace=True)
        grade = pd.DataFrame(columns=shoot_df.columns, index=shoot_df.index)

        def grade_assess(subject):
            for i in shoot_df.index:
                if result_single['总分'].at[i] != 0:
                    if (result_single[subject].at[i]) >= (result_single['总分'].at[i]) * 0.8:
                        if (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'A'
                        elif (contribution_df[subject].at[i] >= 0.7) & (shoot_df[subject].at[i] < 0.6):
                            grade[subject].at[i] = 'B'
                        elif (contribution_df[subject].at[i] < 0.7) & (shoot_df[subject].at[i] >= 0.6):
                            grade[subject].at[i] = 'C'
                        else:
                            grade[subject].at[i] = 'D'
                    else:
                        grade[subject].at[i] = 'E'
                else:
                    grade[subject].at[i] = 'F'

        grade_assess('语文')
        grade_assess('数学')
        grade_assess('英语')
        grade_assess('物理')
        grade_assess('化学')
        grade_assess('生物')
        shoot_df = shoot_df.applymap(lambda x: format(x, '.2%'))
        contribution_df = contribution_df.applymap(lambda x: format(x, '.2%'))
        final_grade = pd.concat([contribution_df, shoot_df, grade],
                                keys=['贡献率', '命中率', '等级'])
        return result_final_science, final_grade, unmatched_students_science

    def good_scores_science_ratio(self, chn, math, eng, pol, his, geo, total):

        single_chn_science, double_chn_science = self.get_single_double_score(self.df_science, '语文', chn, total)
        single_math_science, double_math_science = self.get_single_double_score(self.df_science, '数学', math, total)
        single_eng_science, double_eng_science = self.get_single_double_score(self.df_science, '英语', eng, total)
        single_phys_science, double_phys_science = self.get_single_double_score(self.df_science, '物理', pol, total)
        single_chem_science, double_chem_science = self.get_single_double_score(self.df_science, '化学', his, total)
        single_bio_science, double_bio_science = self.get_single_double_score(self.df_science, '生物', geo, total)
        single_total_science, double_total_science = self.get_single_double_score(self.df_science, '总分', total, total)

        name_num = self.df_science.groupby(['班级'])['姓名'].count()
        name_num.name = '参考人数'

        result_single = pd.concat([name_num, single_chn_science, single_math_science, single_eng_science,
                                   single_phys_science, single_chem_science, single_bio_science, single_total_science],
                                  axis=1)
        result_double = pd.concat(
            [name_num, double_chn_science, double_math_science, double_eng_science, double_phys_science,
             double_chem_science, double_bio_science, double_total_science], axis=1)

        result_single.loc['理科共计'] = [result_single['参考人数'].sum(),
                                     result_single['语文'].sum(),
                                     result_single['数学'].sum(),
                                     result_single['英语'].sum(),
                                     result_single['物理'].sum(),
                                     result_single['化学'].sum(),
                                     result_single['生物'].sum(),
                                     result_single['总分'].sum()
                                     ]

        result_single = self.good_scores_ratio(result_single, '总分')
        result_single = self.good_scores_ratio(result_single, '语文')
        result_single = self.good_scores_ratio(result_single, '数学')
        result_single = self.good_scores_ratio(result_single, '英语')
        result_single = self.good_scores_ratio(result_single, '物理')
        result_single = self.good_scores_ratio(result_single, '化学')
        result_single = self.good_scores_ratio(result_single, '生物')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '物理', '物理上线率', '化学', '化学上线率', '生物', '生物上线率', '总分', '总分上线率']
        result_single = result_single[order]

        result_double.loc['理科共计'] = [result_double['参考人数'].sum(),
                                     result_double['语文'].sum(),
                                     result_double['数学'].sum(),
                                     result_double['英语'].sum(),
                                     result_double['物理'].sum(),
                                     result_double['化学'].sum(),
                                     result_double['生物'].sum(),
                                     result_double['总分'].sum()]
        result_double = self.good_scores_ratio(result_double, '总分')
        result_double = self.good_scores_ratio(result_double, '语文')
        result_double = self.good_scores_ratio(result_double, '数学')
        result_double = self.good_scores_ratio(result_double, '英语')
        result_double = self.good_scores_ratio(result_double, '物理')
        result_double = self.good_scores_ratio(result_double, '化学')
        result_double = self.good_scores_ratio(result_double, '生物')
        order = ['参考人数', '语文', '语文上线率', '数学', '数学上线率', '英语', '英语上线率',
                 '物理', '物理上线率', '化学', '化学上线率', '生物', '生物上线率', '总分', '总分上线率']
        result_double = result_double[order]
        single_double_concat_science = pd.concat([result_single, result_double], keys=['单有效', '双有效'])

        return single_double_concat_science

    def get_unmatched_arts(self, chn, math, eng, pol, his, geo, total):
        df_chn = self.get_unmatched_students(self.df_arts, '语文', chn, total)
        df_math = self.get_unmatched_students(self.df_arts, '数学', math, total)
        df_eng = self.get_unmatched_students(self.df_arts, '英语', eng, total)
        df_pol = self.get_unmatched_students(self.df_arts, '政治', pol, total)
        df_his = self.get_unmatched_students(self.df_arts, '历史', his, total)
        df_geo = self.get_unmatched_students(self.df_arts, '地理', geo, total)
        df_unmatched = pd.concat([df_chn, df_math, df_eng, df_pol, df_his, df_geo], axis=1)

        df_cnh_num = df_chn.groupby('班级')[['语文']].count()
        df_math_num = df_math.groupby('班级')[['数学']].count()
        df_eng_num = df_eng.groupby('班级')[['英语']].count()
        df_pol_num = df_pol.groupby('班级')[['政治']].count()
        df_his_num = df_his.groupby('班级')[['历史']].count()
        df_geo_num = df_geo.groupby('班级')[['地理']].count()
        df_unmatched_num = pd.concat([df_cnh_num, df_math_num, df_eng_num, df_pol_num, df_his_num, df_geo_num], axis=1)

        return df_unmatched, df_unmatched_num

    def line_betweens(self, total=None, total_science=None):

        line_condition = (self.df_arts['总分'] >= total - 20) & (self.df_arts['总分'] <= total + 20)
        line_condition_science = (self.df_science['总分'] >= total_science - 20) & (
                self.df_science['总分'] <= total_science + 20)
        df_line_arts = self.df_arts.loc[line_condition, :]
        df_line_science = self.df_science.loc[line_condition_science, :]
        writer = pd.ExcelWriter(r'D:\成绩统计结果\本次考试踩线生分班名单.xlsx')
        class_num = list(df_line_arts['班级'].drop_duplicates())
        class_num_science = list(df_line_science['班级'].drop_duplicates())
        for i in class_num:
            class_name = df_line_arts[df_line_arts['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name = class_name.loc[:, ['序号', '姓名', '班级', '语文', '数学', '英语',
                                            '政治', '历史', '地理', '总分', '排名']]
            class_name.to_excel(writer, sheet_name=i, index=False)

        for i in class_num_science:
            class_name = df_line_science[df_line_science['班级'] == i].reset_index(drop=True)
            class_name['序号'] = [k + 1 for k in class_name.index]
            class_name = class_name.loc[:, ['序号', '姓名', '班级', '语文', '数学', '英语',
                                            '物理', '化学', '生物', '总分', '排名']]
            class_name.to_excel(writer, sheet_name=i, index=False)

        writer.close()

    def class_divided(self):
        """
        计算获得文理科各班成绩表
        :return:
        """

        self.class_rank()
        class_No_arts = list(self.df_arts['班级'].drop_duplicates())
        class_NO_science = list(self.df_science['班级'].drop_duplicates())
        writer = pd.ExcelWriter(r'D:\成绩统计结果\本次考试文理各班成绩表.xlsx')
        for i in class_No_arts:
            class_arts = self.df_arts[self.df_arts['班级'] == i].reset_index(drop=True)
            class_arts['序号'] = [k + 1 for k in class_arts.index]
            class_arts['综合'] = class_arts['政治'] + class_arts['历史'] + class_arts['地理']
            class_arts = class_arts.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '综合',
                                            '政治', '历史', '地理', '总分', '排名']]
            class_arts.to_excel(writer, sheet_name=i, index=False)
        for i in class_NO_science:
            class_science = self.df_science[self.df_science['班级'] == i].reset_index(drop=True)
            class_science['序号'] = [k + 1 for k in class_science.index]
            class_science['综合'] = class_science['物理'] + class_science['化学'] + class_science['生物']
            class_science = class_science.loc[:, ['序号', '班级', '姓名', '语文', '数学', '英语', '综合',
                                                  '物理', '化学', '生物', '总分', '排名']]
            class_science.to_excel(writer, sheet_name=i, index=False)
        writer.close()

    def top_class_student(self):
        """
        计算各班单科前N名
        """
        class_names = self.df_arts['班级'].unique()
        class_names_science = self.df_science['班级'].unique()
        writer = pd.ExcelWriter(r'D:\成绩统计结果\各班前N名.xlsx')
        for i in class_names:
            class_name = self.df_arts[self.df_arts['班级'] == i]
            class_name = class_name.copy()
            top_chn = self.get_top_n(class_name, '语文', n=ScoreAnalysis.top_n)
            top_math = self.get_top_n(class_name, '数学', n=ScoreAnalysis.top_n)
            top_eng = self.get_top_n(class_name, '英语', n=ScoreAnalysis.top_n)
            top_pol = self.get_top_n(class_name, '政治', n=ScoreAnalysis.top_n)
            top_his = self.get_top_n(class_name, '历史', n=ScoreAnalysis.top_n)
            top_geo = self.get_top_n(class_name, '地理', n=ScoreAnalysis.top_n)
            top_total = self.get_top_n(class_name, '总分', n=ScoreAnalysis.top_n)
            class_top = pd.concat([top_chn, top_math, top_eng, top_pol, top_his, top_geo, top_total], axis=1)
            class_top.index = [i + 1 for i in class_top.index]
            class_top.drop(columns='班级', inplace=True)
            class_top.to_excel(writer, sheet_name=i, index_label='序号')
        for i in class_names_science:
            class_name = self.df_science[self.df_science['班级'] == i]
            class_name = class_name.copy()
            top_chn = self.get_top_n(class_name, '语文', n=ScoreAnalysis.top_n)
            top_math = self.get_top_n(class_name, '数学', n=ScoreAnalysis.top_n)
            top_eng = self.get_top_n(class_name, '英语', n=ScoreAnalysis.top_n)
            top_pol = self.get_top_n(class_name, '物理', n=ScoreAnalysis.top_n)
            top_his = self.get_top_n(class_name, '化学', n=ScoreAnalysis.top_n)
            top_geo = self.get_top_n(class_name, '生物', n=ScoreAnalysis.top_n)
            top_total = self.get_top_n(class_name, '总分', n=ScoreAnalysis.top_n)
            class_top = pd.concat([top_chn, top_math, top_eng, top_pol, top_his, top_geo, top_total], axis=1)
            class_top.index = [i + 1 for i in class_top.index]
            class_top.drop(columns='班级', inplace=True)
            class_top.to_excel(writer, sheet_name=i, index_label='序号')

        writer.close()

    def class_rank(self):
        """
        计算文理科学生排名
        :return:
        """
        self.df_arts['排名'] = self.df_arts['总分'].rank(method='min', ascending=False)
        # self.df_arts['排名'] = self.df_arts['排名'].apply(lambda x: format(int(x)))
        self.df_arts.sort_values(by='总分', ascending=False, inplace=True)

        self.df_science['排名'] = self.df_science['总分'].rank(method='min', ascending=False)
        # self.df_science['排名'] = self.df_science['排名'].apply(lambda x: format(int(x)))
        self.df_science.sort_values(by='总分', ascending=False, inplace=True)

    def score_label(self):
        """
        计算打印考生个人成绩单
        """

        self.class_rank()
        exam_arts = self.df_arts.loc[:, ['班级', '姓名', '语文', '数学', '英语', '政治', '历史', '地理', '总分', '排名']]
        exam_science = self.df_science.loc[:, ['班级', '姓名', '语文', '数学', '英语', '物理', '化学', '生物', '总分', '排名']]
        exam_arts.sort_values(by=['班级', '总分'], inplace=True, ascending=[True, False], ignore_index=True)
        exam_science.sort_values(by=['班级', '总分'], inplace=True, ascending=[True, False], ignore_index=True)

        for i in exam_arts.index:
            exam_arts.loc[i + 0.5] = exam_arts.columns
        exam_arts.sort_index(inplace=True, ignore_index=True)
        for i in exam_science.index:
            exam_science.loc[i + 0.5] = exam_science.columns
        exam_science.sort_index(inplace=True)
        with pd.ExcelWriter(r'D:\成绩统计结果\本次考试学生个人成绩单.xlsx') as writer:
            exam_arts.to_excel(writer, sheet_name='文科成绩单', index=False)
            exam_science.to_excel(writer, sheet_name='理科成绩单', index=False)

    def combine_files(self, exam_record=r'D:\成绩统计结果\本次考试成绩分析统计.xlsx'):
        """
        计算各类统考相关数据，有效分数据由市上或区上统计获得。调用前，需先输入有效分。
        """

        av_arts, av_science = self.get_av()
        self.class_divided()
        self.line_betweens(total=ScoreAnalysis.arts_scores[-1], total_science=ScoreAnalysis.science_scores[-1])

        arts, grades_arts, unmatched_arts = self.goodscore_arts(*ScoreAnalysis.arts_scores)
        science, grades_science, unmatched_science = self.goodscore_science(*ScoreAnalysis.science_scores)
        with pd.ExcelWriter(exam_record) as writer:
            self.df_arts.to_excel(writer, sheet_name='文科总表')
            self.df_science.to_excel(writer, sheet_name='理科总表')
            av_arts.to_excel(writer, sheet_name='文科平均分', float_format='%.2f')

            av_science.to_excel(writer, sheet_name='理科平均分', float_format='%.2f')
            arts.to_excel(writer, sheet_name='文科有效分')
            unmatched_arts.to_excel(writer, sheet_name='文科错位生', index=False)
            science.to_excel(writer, sheet_name='理科有效分')
            unmatched_science.to_excel(writer, sheet_name='理科错位生', index=False)
            grades_arts.to_excel(writer, sheet_name='文科贡献率', float_format='%.2f')
            grades_science.to_excel(writer, sheet_name='理科贡献率', float_format='%.2f')
        print('successfully done')

    def combine_files_school(self, exam_record=r'D:\成绩统计结果\本次考试成绩分析统计.xlsx', goodtotal_arts=None,
                             goodtotal_science=None):

        """
        计算学校考试相关数据，平均分，有效分，分班成绩表等。

        """

        chn, math, eng, pol, his, geo, total = self.get_goodscores_arts(goodtotal_arts)
        chn_science, math_science, eng_science, phys, chem, bio, total_science = self.get_goodscores_science(
            goodtotal_science)

        self.class_divided()
        self.line_betweens(total=total, total_science=total_science)
        av_arts, av_science = self.get_av()
        arts, grades_arts, unmatched_arts = self.goodscore_arts(chn, math, eng, pol, his, geo, total)
        science, grades_science, unmatched_science = self.goodscore_science(chn_science, math_science, eng_science,
                                                                            phys, chem, bio,
                                                                            total_science)
        with pd.ExcelWriter(exam_record) as writer:
            self.df_arts.to_excel(writer, sheet_name='文科总表')
            self.df_science.to_excel(writer, sheet_name='理科总表')
            av_arts.to_excel(writer, sheet_name='文科平均分', float_format='%.2f')
            av_science.to_excel(writer, sheet_name='理科平均分', float_format='%.2f')
            arts.to_excel(writer, sheet_name='文科有效分')
            unmatched_arts.to_excel(writer, sheet_name='文科错位生', index=False)
            science.to_excel(writer, sheet_name='理科有效分')
            unmatched_science.to_excel(writer, sheet_name='理科错位生', index=False)
            grades_arts.to_excel(writer, sheet_name='文科贡献率', float_format='%.2f')
            grades_science.to_excel(writer, sheet_name='理科贡献率', float_format='%.2f')
        print('successfully done')

    def arts_science_combined(self):
        arts_av, science_av = self.get_av()
        arts, grades_arts, unmatched_arts = self.goodscore_arts(*ScoreAnalysis.arts_scores)
        science, grades_science, unmatched_science = self.goodscore_science(*ScoreAnalysis.science_scores)
        arts_percentage = self.good_scores_arts_ratio(*ScoreAnalysis.arts_scores)
        science_percentage = self.good_scores_science_ratio(*ScoreAnalysis.science_scores)
        arts_av = self.write_open(arts_av)
        science_av = self.write_open(science_av)
        arts = self.write_open(arts)
        science = self.write_open(science)
        grades_arts = self.write_open(grades_arts)
        grades_science = self.write_open(grades_science)
        arts_percentage = self.write_open(arts_percentage)
        science_percentage = self.write_open(science_percentage)
        arts_science_av = pd.concat([arts_av, science_av])
        arts_science_goodscores = pd.concat([arts, science], ignore_index=True)
        arts_science_grade = pd.concat([grades_arts, grades_science], ignore_index=True)
        arts_science_percentage = pd.concat([arts_percentage, science_percentage], ignore_index=True)
        with pd.ExcelWriter(r'D:\成绩统计结果\文理有效分统计分析.xlsx') as writer:
            arts_science_av.to_excel(writer, sheet_name='文理平均分', float_format='%.2f', index=False)
            arts_science_goodscores.to_excel(writer, sheet_name='文理有效分', index=False)
            arts_science_grade.to_excel(writer, sheet_name='文理等级评定', float_format='%.2f', index=False)
            arts_science_percentage.to_excel(writer, sheet_name='文理有效百分比', index=False)

    def arts_science_combined_school(self, goodtotal_arts=None, goodtotal_science=None):
        chn, math, eng, pol, his, geo, total = self.get_goodscores_arts(goodtotal_arts)
        chn_science, math_science, eng_science, phys, chem, bio, total_science = self.get_goodscores_science(
            goodtotal_science)
        arts_av, science_av = self.get_av()

        arts, grades_arts, unmatched_arts = self.goodscore_arts(chn, math, eng, pol, his, geo, total)
        science, grades_science, unmatched_science = self.goodscore_science(chn_science, math_science, eng_science,
                                                                            phys, chem, bio,
                                                                            total_science)
        arts_percentage = self.good_scores_arts_ratio(chn, math, eng, pol, his, geo, total)
        science_percentage = self.good_scores_science_ratio(chn_science, math_science, eng_science,
                                                            phys, chem, bio, total_science)

        arts_av = self.write_open(arts_av)
        science_av = self.write_open(science_av)
        arts = self.write_open(arts)
        science = self.write_open(science)
        grades_arts = self.write_open(grades_arts)
        grades_science = self.write_open(grades_science)
        arts_percentage = self.write_open(arts_percentage)
        science_percentage = self.write_open(science_percentage)
        arts_science_av = pd.concat([arts_av, science_av])
        arts_science_goodscores = pd.concat([arts, science])
        arts_science_grade = pd.concat([grades_arts, grades_science])
        arts_science_percentage = pd.concat([arts_percentage, science_percentage], ignore_index=True)
        with pd.ExcelWriter(r'D:\成绩统计结果\文理有效分统计分析.xlsx') as writer:
            arts_science_av.to_excel(writer, sheet_name='文理平均分', float_format='%.2f', index=False)
            arts_science_goodscores.to_excel(writer, sheet_name='文理有效分', index=False)
            arts_science_grade.to_excel(writer, sheet_name='文理等级评定', float_format='%.2f', index=False)
            arts_science_percentage.to_excel(writer, sheet_name='文理有效百分比', index=False)

    @staticmethod
    def get_student_number_class(df_data):
        """
        get the numbers of the students in each class
        :param df_data: the df of df table
        :return: the number of students
        """
        student_number_class = df_data['班级'].value_counts()
        student_number_class.name = '参考人数'
        student_number_class['年级'] = student_number_class.sum()
        return student_number_class

    @staticmethod
    def get_subject_good_score(data, subject, total):
        """
        获取各科有效分
        :param data: df数据
        :param subject: 学科名
        :param total: 上线总分
        :return: 学科有效分
        """
        good_score_data = data.loc[data['总分'] >= total]
        subject_av = good_score_data[subject].mean()
        total_av = good_score_data['总分'].mean()
        subject_good_score = round(subject_av * total / total_av)
        return subject_good_score

    # @staticmethod
    # def get_top_n(df_data, subject, n=None):
    #     """
    #     get top n students in an exam
    #     :param df_data:
    #     :param subject:
    #     :param n:
    #     :return:
    #     """
    #     df = df_data.sort_values(by=subject, ascending=False).reset_index()
    #     top_students = df.loc[0:n - 1, ['班级', '姓名', subject]]
    #     while df[subject].at[n] == df[subject].at[n - 1]:
    #         top_students = top_students.append(df.loc[n, ['班级', '姓名', subject]])
    #         n += 1
    #     return top_students
    @staticmethod
    def get_top_n(df_data, subject, n=None):
        df_data.sort_values(by=subject, ascending=False, inplace=True)
        df_data.reset_index(inplace=True, drop=True)
        top_students = df_data.loc[0:n - 1, ['班级', '姓名', subject]]
        top_list = []
        while df_data[subject].at[n] == df_data[subject].at[n - 1]:
            # top_students = top_students.append(df_data.loc[n, ['班级', '姓名', subject]])
            # 新版pandas取消了appedn方法，上一行代码改为：
            top_list.append(df_data.loc[n, ['班级', '姓名', subject]])
            n += 1
        top_list_df = pd.DataFrame(top_list, columns=top_students.columns)
        top_students = pd.concat([top_students, top_list_df])

        return top_students

    @staticmethod
    def get_single_double_score(data, subject, subject_score, total_score):
        """
        get the good scores in an exam
        :param data:
        :param subject:
        :param subject_score:
        :param total_score:
        :return:
        """
        single = data[data[subject] >= subject_score].groupby(['班级'])[subject].count()
        data_double = data[data['总分'] >= total_score]
        double = data_double[data_double[subject] >= subject_score].groupby(['班级'])[subject].count()
        return single, double

    @staticmethod
    def get_unmatched_students(data, subject, subject_score, total_score):
        """
        get the unmatched students in an exam
        :param data:
        :param subject:
        :param subject_score:
        :param total_score:
        :return:
        """
        df2 = data[data['总分'] >= total_score]
        df_unmatched = df2.loc[:, ['班级', '姓名', subject]].loc[df2[subject] < subject_score].sort_values(
            by=['班级', subject], ascending=[True, False]).reset_index(drop=True)
        return df_unmatched

    @staticmethod
    def good_scores_ratio(data, subject):
        """
        get the ratio of a subject that is above the good total.
        :param data:
        :param subject:
        :return:
        """
        data[subject + '上线率'] = data[subject] / data['参考人数']
        data[subject + '上线率'] = data[subject + '上线率'].apply(lambda x: format(x, '.1%'))
        return data

    @staticmethod
    def write_open(df_data):
        """
        used to concat different dataframes.
        :param df_data:
        :return:
        """
        df_data.to_excel('temp_data.xlsx')
        df_new = pd.read_excel('temp_data.xlsx', header=None)
        os.remove('temp_data.xlsx')
        return df_new

    @staticmethod
    def make_directory(f):
        def wrapper(*args, **kwargs):
            if not os.path.exists('D:\\成绩统计结果'):
                os.makedirs('D:\\成绩统计结果')
            result = f(*args, **kwargs)
            return result

        return wrapper

    @staticmethod
    def use_time(f):
        def wrapper(*args, **kwargs):
            t1 = time.time()
            results = f(*args, **kwargs)
            t2 = time.time()
            print(f'主程序{f.__name__}用时{(t2 - t1):.1f}秒')
            return results

        return wrapper

    @staticmethod
    def title_lines(f):
        def wrapper(*args, **kwargs):
            print(f'+++++++++++++++高中考试成绩分析处理+++++++++++++++')
            f(*args, **kwargs)
            print(f'--------------成绩分析处理完毕，程序已关闭-----------------')

        return wrapper

    def show_menu(self):
        print(self.__str__())

        while True:
            flag = eval(input(f'按键功能选择:\n '
                              f'    1:年级考试成绩分析；\n '
                              f'    2:区级以上考试成绩分析；\n'
                              f'     3:生成考室安排表；\n '
                              f'    4:生成成绩单；\n '
                              f'    5:生成与上次考试对比表，\n'
                              f'     6:生成各科前N名;\n'
                              f'     7:生成各班前N名\n'
                              f'     8:生成平均分占比\n'
                              f'     9:按其它数字键退出程序.\n请选择：'))
            if flag == 1:
                arts = eval(input('请输入本次考试文科上线总分：'))
                science = eval(input('请输入本次考试理科上线总分：'))
                score_arts = self.get_goodscores_arts(goodtotal_arts=arts)
                print(score_arts)
                score_science = self.get_goodscores_science(goodtotal_science=science)
                print(score_science)
                self.combine_files_school(goodtotal_arts=arts, goodtotal_science=science)
                self.arts_science_combined_school(goodtotal_arts=arts, goodtotal_science=science)
                print('成绩分析已完成，谢谢使用！')
                break
            elif flag == 2:
                # ScoreAnalysis.arts_scores = [int(i) for i in input('请输入文科有效分及上线总分，以空格隔开:').split()]
                # ScoreAnalysis.science_scores = [int(i) for i in input('请输入理科有效分及上线总分，以空格隔开:').split()]
                self.combine_files()
                self.arts_science_combined()
                print('成绩分析已完成，谢谢使用！')
                break
            elif flag == 3:
                ScoreAnalysis.numbers_in_room = eval(input('请输入每间考室的人数：'))
                self.exam_room_info()
                print('考室信息已生成，谢谢使用！')
                break
            elif flag == 4:
                self.score_label()
                print('学生个人成绩单已生成，谢谢使用！')
                break

            elif flag == 5:
                exam1 = input('请输入本次考试名称：')
                exam2 = input('请输入上次考试名称：')
                ContrastScores.contrast(exam1=exam1, exam2=exam2)
                print('对比成绩已生成，谢谢使用！')
                break
            elif flag == 6:
                ScoreAnalysis.top_n = eval(input('请输入要计算的前N名：'))
                self.top_n_students()
                break
            elif flag == 7:
                ScoreAnalysis.top_n = eval(input('请输入要计算的前N名：'))
                self.top_class_student()
                break
            elif flag == 8:
                self.get_av_diagram()
                break
            else:
                break


class RankDistribution(object):

    def __init__(self, file_path, file_path_1):
        self.file_path = file_path
        self.file_path_1 = file_path_1
        # file_path = r'D:\work documents\高2021级\高一下半期考试\高2021级高一（下）半期考试.xlsx'
        # file_path_1 = r'D:\work documents\高2021级\高一下半期考试\高2021级文理分班信息表(对比基准）.xlsx'
        self.df_arts = pd.read_excel(file_path, sheet_name='文科', index_col='序号')
        self.df_science = pd.read_excel(file_path, sheet_name='理科', index_col='序号')
        self.df_arts_1 = pd.read_excel(file_path_1, sheet_name='文科', index_col='序号')
        self.df_science_1 = pd.read_excel(file_path_1, sheet_name='理科', index_col='序号')

    def rank_by_class_loop(self, n, arts_data=None, science_data=None):
        arts_subjects = ['语文', '数学', '英语', '政治', '历史', '地理', '总分']
        science_subjects = ['语文', '数学', '英语', '物理', '化学', '生物', '总分']
        arts_subjects_range = []
        science_subjects_range = []
        for subject in arts_subjects:
            subject_range = self.rank_by_subject(arts_data, subject, n)
            arts_subjects_range.append(subject_range)
        arts_range = pd.concat(arts_subjects_range, axis=1)
        arts_range.index.rename('班级', inplace=True)
        arts_range.sort_index(inplace=True, ascending=True)
        arts_range.fillna(0, inplace=True)

        # print(arts_range)
        for subject in science_subjects:
            subject_range = self.rank_by_subject(science_data, subject, n)
            science_subjects_range.append(subject_range)
        science_range = pd.concat(science_subjects_range, axis=1)
        science_range.index.rename('班级', inplace=True)
        science_range.sort_index(inplace=True, ascending=True)
        science_range.fillna(0, inplace=True)

        return arts_range, science_range

    def contrast_range(self, n):
        arts_range, science_range = self.rank_by_class_loop(n, arts_data=self.df_arts, science_data=self.df_science)

        arts_contrast, science_contrast = self.rank_by_class_loop(n, arts_data=self.df_arts_1,
                                                                  science_data=self.df_science_1)

        arts_contrast_result = arts_range - arts_contrast
        art_result = pd.concat([arts_range, arts_contrast, arts_contrast_result], keys=['分段人数', '对比人数', '对比变化'])
        science_contrast_result = science_range - science_contrast
        science_result = pd.concat([science_range, science_contrast, science_contrast_result],
                                   keys=['分段人数', '对比人数', '对比变化'])
        # print(art_result)

        return art_result, science_result

    def main_rank_range(self):
        arts_rank_10, science_rank_10 = self.rank_by_class_loop(10, arts_data=self.df_arts,
                                                                science_data=self.df_science)
        arts_rank_30, science_rank_30 = self.rank_by_class_loop(30, arts_data=self.df_arts,
                                                                science_data=self.df_science)
        arts_rank_50, science_rank_50 = self.rank_by_class_loop(50, arts_data=self.df_arts,
                                                                science_data=self.df_science)
        arts_rank_100, science_rank_100 = self.rank_by_class_loop(100, arts_data=self.df_arts,
                                                                  science_data=self.df_science)

        arts_rank_10 = self.write_open(arts_rank_10)
        science_rank_10 = self.write_open(science_rank_10)
        arts_rank_30 = self.write_open(arts_rank_30)
        science_rank_30 = self.write_open(science_rank_30)
        arts_rank_50 = self.write_open(arts_rank_50)
        science_rank_50 = self.write_open(science_rank_50)
        arts_rank_100 = self.write_open(arts_rank_100)
        science_rank_100 = self.write_open(science_rank_100)

        rank_range_arts = pd.concat([arts_rank_10, arts_rank_30, arts_rank_50, arts_rank_100])
        rank_range_science = pd.concat([science_rank_10, science_rank_30, science_rank_50, science_rank_100])
        rank_range_arts.fillna(0, inplace=True)
        rank_range_science.fillna(0, inplace=True)
        with pd.ExcelWriter(r'D:\成绩统计结果\名次段分布.xlsx') as writer:
            rank_range_arts.to_excel(writer, sheet_name='文科名次分布', index=False)
            rank_range_science.to_excel(writer, sheet_name='理科名次分布', index=False)

        print(f'successfully done')

    def run_rank_change(self):
        arts_minus_10, science_minus_10 = self.contrast_range(10)
        arts_minus_30, science_minus_30 = self.contrast_range(30)
        arts_minus_50, science_minus_50 = self.contrast_range(50)
        arts_minus_100, science_minus_100 = self.contrast_range(100)
        arts_minus_10 = self.write_open(arts_minus_10)
        arts_minus_30 = self.write_open(arts_minus_30)
        arts_minus_50 = self.write_open(arts_minus_50)
        arts_minus_100 = self.write_open(arts_minus_100)

        science_minus_10 = self.write_open(science_minus_10)
        science_minus_30 = self.write_open(science_minus_30)
        science_minus_50 = self.write_open(science_minus_50)
        science_minus_100 = self.write_open(science_minus_100)
        arts = pd.concat([arts_minus_10, arts_minus_30, arts_minus_50, arts_minus_100], ignore_index=True)
        science = pd.concat([science_minus_10, science_minus_30, science_minus_50, science_minus_100],
                            ignore_index=True)
        # arts.fillna(0, inplace=True)
        # science.fillna(0, inplace=True)
        with pd.ExcelWriter(r'D:\成绩统计结果\名次段对比变化.xlsx') as writer:
            arts.to_excel(writer, sheet_name='文科对比', index=False)
            science.to_excel(writer, sheet_name='理科对比', index=False)
        print('successfully done')

    def main(self):
        self.run_rank_change()
        self.main_rank_range()

    @staticmethod
    def write_open(df_data):
        """
        used to concat different dataframes.
        :param df_data:
        :return:
        """
        df_data.to_excel('temp_data.xlsx')
        df_new = pd.read_excel('temp_data.xlsx', header=None)
        os.remove('temp_data.xlsx')
        return df_new

    @staticmethod
    def rank_by_subject(data, subject, n):
        data['排名'] = data[subject].rank(method='min', ascending=False)
        df_rank = data[data['排名'] <= n]
        df_rank_num = df_rank['班级'].value_counts()
        df_rank_num.name = f'{subject}前{n}名'
        return df_rank_num


if __name__ == '__main__':
    filepath = r'D:\work documents\高2021级\高一下半期考试\高2021级高一（下）半期考试.xlsx'
    ScoreAnalysis.arts_scores = (94, 59, 74, 54, 55, 66, 390)
    ScoreAnalysis.science_scores = (87, 62, 66, 39, 37, 39, 330)
    ScoreAnalysis.top_n = 3
    # ExamRoom.numbers_in_room = 35
    exam_processed = ScoreAnalysis(filepath)
    ContrastScores.path_one = filepath
    ContrastScores.path_two = r'D:\work documents\高2020级\高2020级半期考试.xlsx'


    @exam_processed.make_directory
    @exam_processed.title_lines
    @exam_processed.use_time
    def main():
        # exam_processed.show_menu()
        exam_processed.get_av_diagram()


    main()
