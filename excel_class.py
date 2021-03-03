import numpy as np
import pandas as pd
import json

import sys

from data_manage import get_date, get_db2_data
from funcs import formulas
from util import is_operation, is_function, get_row_column, get_list_row, check_value, ret_row_column, \
    percentile_tu_float, row_list

try:
    import cPickle as pickle
except ImportError:
    import pickle


class ExcelMain():
    temp_data = {}
    data = None
    db2 = None
    func_dic = {}
    return_data = {}

    def __init__(self, params=None):
        self.data = get_date()
        self.get_all_caibao_func()
        # self.temp_data = self.get_all_global_data()
        self.db2 = get_db2_data()

        # params = {
        #     # 'province': province,
        #     # 'company': company,
        #     # 'pool': pool,
        #     # 'fannum': fannum,
        #     # 'setline_len': float(setline_len or 0),
        #     # 'pit_road': float(pit_road or 0),
        #     # 'financing_way': financing_way,
        #     # 'long_term_financingrate': float(long_term_financingrate or 0),
        #     # 'deadline': float(deadline or 0),
        #     # 'mode_repayment': mode_repayment,
        #     # 'terrain': terrain,
        #     # 'unit_price': float(unit_price),
        #     # 'tower_price': float(tower_price or 0),
        #     # 'anchor_price': float(anchor_price),
        #     # 'nbooster': u"110kV升压站" if nbooster else u"110kV升压站",
        #     # 110kV升压站 220kV升压站 110升压站扩建
        #     # 'ten_switching': float(ten_switching),
        #     # 'thirtyfive_switching': float(thirtyfive_switching),
        #     # 'total_captical': total_captical,
        #     # 'construction_time': int(construction_time),
        #     # 'cut_hour': float(cut_hour),
        #     # 'power_limiting_year': float(power_limiting_year),
        #     # 'power_limiting_rate': float(power_limiting_rate),
        #     # 'market_power_ratio': float(market_power_ratio),
        #     # 'market_power_loss': float(market_power_loss),
        #     # 'market_year': float(market_year),
        #     # 'base_reinforcement': base_reinforcement,
        #     # 'total_concreteprice': total_concreteprice
        # }

        try:
            types = params["pool"][0] or "GW131/2300-90A"
            turbine_number = params["fannum"][0]
            turbine_price = params["unit_price"][0]
        except:
            types = "GW131/2300-90A"
            turbine_number = 0
            turbine_price = 0
        try:
            types2 = params["pool"][1]
            turbine_number2 = params["fannum"][1]
            turbine_price2 = params["unit_price"][1]
        except:
            types2 = ""
            turbine_number2 = 0
            turbine_price2 = 0
        try:
            types3 = params["pool"][2]
            turbine_number3 = params["fannum"][2]
            turbine_price3 = params["unit_price"][2]
        except:
            types3 = ""
            turbine_number3 = 0
            turbine_price3 = 0
        # 折减后等效小时数
        cut_hour = params['cut_hour']

        terrain = params['terrain'] or "平原"
        is_keyan = "否"
        province = params['province']
        company = params['company']
        design_capacity = params['total_captical']
        # 塔筒价格
        tower_price = params['tower_price']
        # 锚栓价格
        anchor_price = params['anchor_price']
        # 基础垫层混凝土（单价）
        jichu_diancheng_hunningt = params['total_concreteprice'] or 550
        # 基础混凝土（单价）
        jichuhunningt = params['base_reinforcement']
        # 基础钢筋（单价）
        gangjin = params['foundation_concrete1']
        # 集电线路长度
        setline_len = params["setline_len"]
        # 场内道路长度
        pit_road = params['pit_road']
        # 10kV开关站
        kw10 = params['ten_switching']
        # 35kV开关站
        kw35 = params['thirtyfive_switching']
        # 融资方式
        financing_way = params['financing_way']
        # 贷款期利率（含税）
        long_term_financingrate = params['long_term_financingrate']
        # 期限
        deadline = params['deadline']
        # 还款方式
        mode_repayment = params['mode_repayment']
        # 送出工程
        send_project = params['send_project']
        # 升压站类型
        nbooster = params['nbooster']
        # 建设工期
        construction_time = params['construction_time']
        # 限电年
        power_limiting_year = params['power_limiting_year']
        # 限电率
        power_limiting_rate = params['power_limiting_rate']
        # 市场交易比例
        market_power_ratio = params['market_power_ratio']
        # 市场交易年限
        market_year = params['market_year']
        # 市场交易电价
        market_power_loss = params['market_power_loss']
        self.set_fan_type_data(types, 1)
        self.set_fan_type_data(types2, 2)
        self.set_fan_type_data(types3, 3)
        self.data["模型"]["E5"] = company
        self.data["模型"]["E6"] = province
        self.data["参数"]['O6'] = province
        self.data["假设表"]['E227'] = province
        self.data["模型"]["E8"] = design_capacity
        self.data["参数"]["O8"] = design_capacity
        self.data["假设表"]["E232"] = design_capacity / 10
        self.data["模型"]["E13"] = power_limiting_year
        if power_limiting_rate:
            self.data["模型"]["E14"] = power_limiting_rate
        self.data["模型"]["E15"] = market_year
        self.data["模型"]["E16"] = market_power_ratio
        self.data["模型"]["E17"] = market_power_loss
        self.data["模型"]["E18"] = financing_way
        # self.data["模型"]["E19"] = deadline
        self.data["模型"]["E21"] = mode_repayment
        self.data["模型"]["E24"] = long_term_financingrate
        self.data["模型"]["E29"] = tower_price
        self.data["模型"]["E30"] = anchor_price
        self.data["模型"]["E31"] = jichu_diancheng_hunningt
        self.data["模型"]["E32"] = jichuhunningt
        self.data["模型"]["E33"] = gangjin
        self.data["模型"]["E35"] = terrain
        self.data["模型"]['E38'] = construction_time
        self.data["模型"]["E39"] = is_keyan
        self.data["模型"]["E66"] = turbine_price
        self.data["模型"]["E67"] = turbine_price2
        self.data["模型"]["E68"] = turbine_price3
        self.data["模型"]["E69"] = cut_hour
        self.data["模型"]['E62'] = turbine_number
        self.data["模型"]['E63'] = turbine_number2
        self.data["模型"]['E64'] = turbine_number3
        self.data["模型"]['E54'] = types
        self.data["模型"]['E55'] = types2
        self.data["模型"]['E56'] = types3
        self.data["参数"]['O9'] = types
        self.data["参数"]['O10'] = types2
        self.data["参数"]['O11'] = types3
        self.data["模型"]['E103'] = setline_len
        self.data["模型"]['E104'] = pit_road

        self.data["参数"]['O10'] = kw10
        self.data["参数"]['O11'] = kw35
        self.data["参数"]['O14'] = self.data["模型"]['E14']
        self.data["参数"]['O39'] = cut_hour
        self.data["参数"]['O38'] = construction_time

        self.data["假设表"]['H40'] = nbooster
        self.data["假设表"]['H41'] = send_project
        self.check_data()

    def check_data(self):
        a = []
        if self.data["财报"]["D34"] == "N":
            a.append(35)
            a.append(36)
            a.append(37)
        if self.data["财报"]["D42"] == "N":
            a.append(43)
            a.append(44)
            a.append(45)
        for j in a:
            for i in get_list_row("NC"):
                self.set_date(i + str(j), 0, "财报")

    def get_all_global_data(self):
        # 获取总表数据
        try:
            with open("data_all_json.txt", "r", encoding="utf-8") as f:
                data = f.read()
                data = json.loads(data)
            return data
        except Exception as e:
            return {}

    def format_parser(self, exp, sheep):
        # 解析非公式类数据
        if ":" in exp:
            raise Exception("error ':'")
        exp = exp.replace("=", "")
        exp_list = list(exp)
        temp = ''
        i = 0
        length = len(exp_list)
        for item in exp_list:
            if is_operation(item):
                exp = self.ret_sp_exp(exp, temp, sheep)
                temp = ''
            else:
                temp += item
            if i == length - 1:
                if temp:
                    exp = self.ret_sp_exp(exp, temp, sheep)
                break
            i += 1
        return eval(exp)

    def set_fan_type_data(self, turbine, number=1):

        df = pd.DataFrame(self.db2)
        val = df.loc[df[1] == str(turbine)]
        if not val.loc[::][0].tolist():
            val = pd.DataFrame([[0] * 28])
        list_come = get_list_row("AG")
        dicts_db2 = {1: "56", 2: "57", 3: "58"}
        # 机型高度
        dicts_gaodu = {1: "58", 2: "59", 3: "60"}
        # 机型价格
        dicts_jaige = {1: "66", 2: "67", 3: "68"}
        # 钢塔-重量
        dicts_gangzhong = {1: "72", 2: '77', 3: '82'}
        # 锚栓-重量
        dicts_maozhong = {1: "73", 2: '78', 3: '83'}
        # 混塔段-总价
        dicts_hunta = {1: "73", 2: '78', 3: '83'}
        # 整机吊安装费（单价）
        dicts_diaozhuang = {1: "75", 2: '80', 3: '85'}
        # 垫层混凝土
        dicts_dchnt = {1: "88", 2: '93', 3: '98'}
        # 混凝土
        dicts_hnt = {1: "89", 2: '94', 3: '99'}
        # 钢筋
        dicts_gj = {1: "90", 2: '95', 3: '100'}

        for i, v in enumerate(val.iloc[0]):
            if i == 1:
                continue
            self.data["DB2"][list_come[i + 1] + dicts_db2[number]] = v

        self.data["模型"]["E" + dicts_gaodu[number]] = val.iloc[0][9]
        # data["模型"]["E" + dicts_jaige[number]] = val.iloc[0][9]
        self.data["模型"]["E" + dicts_gangzhong[number]] = val.iloc[0][11]
        self.data["模型"]["E" + dicts_maozhong[number]] = val.iloc[0][22]
        self.data["模型"]["E" + dicts_diaozhuang[number]] = val.iloc[0][26]
        self.data["模型"]["E" + dicts_dchnt[number]] = val.iloc[0][16]
        self.data["模型"]["E" + dicts_hnt[number]] = val.iloc[0][17]
        self.data["模型"]["E" + dicts_gj[number]] = val.iloc[0][18]

    def ret_sp_exp(self, exp, temp, sheep):
        # 判断当前分解步骤中是否存在数值
        if is_function(temp):
            raise Exception("存在函数不能解析")
        try:
            temp = float(temp)
        except:
            if isinstance(temp, str) and temp:
                temp2 = temp.replace("$", "")
                new_value = temp2
                if "!" in temp2:
                    sheep, new_value = temp2.split("!")
                value = self.get_temp_data(new_value, sheep)
                if isinstance(value, int) and value == -999999:
                    value = self.get_date_value(new_value, sheep)
                    value = self.eng(value, sheep, new_value)
                exp = exp.replace(temp, str(value))
        return exp

        # 获取临时数据

    def get_temp_data(self, name, sheet):

        data_sheet = self.temp_data.setdefault(sheet, {})
        return data_sheet.setdefault(name, -999999)

    def set_temp_data(self, name, value, sheet):
        # 存储临时数据
        # try:
        self.temp_data[sheet][name] = value
        # except KeyError:
        #     self.temp_data[sheet] = {}
        # self.temp_data[sheet][name] = value

    def get_date_value(self, value: str, sheet: str):
        # 获取数据
        return self.data[sheet].setdefault(value, None)

    def set_date(self, name, value, sheet):

        self.data[sheet][name] = value

    def input_func(self, sheet=None, reference=None, input=None, old_value=None):
        # 返回input输入参数函数
        if reference and input == reference:
            # 遇见OFFSET时设置标志位跳过当前位
            return reference
        if "!" in input:
            # 遇见表跳转解析 ’！‘
            new_sheet = input.split("!")[0]
            new_sheet = new_sheet.replace("'", "")
            input = input.split("!")[1]
        else:
            # 赋值表名不变
            new_sheet = sheet

        if ":" in input:
            sum_sss = self.get_temp_data(input, new_sheet)
            if sum_sss == -999999:
                # 创建循环
                start = input.split(":")[0]
                end = input.split(":")[1]
                # 解析循环的开始和结束
                start_column, start_row = get_row_column(start)
                end_column, end_row = get_row_column(end, "")
                start_column_list = get_list_row(start_column)
                end_column_list = get_list_row(end_column)
                sum_sss = []
                for i in range(start_row, end_row + 1):
                    rows = []
                    for j in [item for item in end_column_list if not item in start_column_list[:-1]]:
                        val = self.get_temp_data(j + str(i), new_sheet)
                        if val == -999999:
                            val = self.eng(self.get_date_value(j + str(i), new_sheet), new_sheet, old_value=j + str(i))
                        if isinstance(val, str) and val.startswith("="):
                            val = self.eng(val, new_sheet, j + str(i))
                        val = check_value(type(val))(val)
                        rows.append(val)
                        # 添加循环出来的每一个值
                        self.set_temp_data(j + str(i), val, new_sheet)
                        self.set_date(j + str(i), val, new_sheet)
                    sum_sss.append(rows)
                if len(sum_sss) == 1:
                    sum_sss = sum_sss[0]
                self.set_temp_data(input, sum_sss, new_sheet)
            sum_sss = np.array(sum_sss)
            return sum_sss
        else:
            val = self.get_temp_data(input, new_sheet)
            if val == -999999 or (isinstance(val, str) and val.startswith("=")):
                # 普通数据时直接去解析坐标的值
                val = self.eng(self.get_date_value(input, new_sheet), new_sheet, input)
                if isinstance(val, str) and val.startswith("="):
                    val = self.eng(val, new_sheet, input)
            val = check_value(type(val))(val)
            self.set_temp_data(input, val, new_sheet)
            self.set_date(input, val, new_sheet)
            return val

    def end_time(self, old_value):
        # 判断 是否是最后工期 大于则不判断直接返回0
        jianshe_time = self.get_temp_data("E230", "假设表")
        if jianshe_time == -999999:
            jianshe_time = self.eng("=E230", "假设表")
        all_time = jianshe_time + 240 + 12 + 7
        source_r, source_c = ret_row_column(old_value)
        if (source_c > all_time and source_r > 5) or (source_r in [507, 508, 519, 520] and source_c not in [20, 6]):
            return True
        else:
            return False

    def eng(self, value: str, sheet: str, old_value: str = "A1", count_func=0):
        # 主要逻辑函数
        # print(sheet + "\t" + old_value + "\t" + str(value) + "\t")
        if isinstance(value, str) and value.startswith("="):
            if sheet == "财报":
                if self.end_time(old_value):
                    return 0

            # 遇见返回行 解析当前坐标并替换函数
            if "COLUMN()" in value:
                r, c = ret_row_column(old_value)
                value = value.replace("COLUMN()", str(c))

            if "%" in value:
                value = percentile_tu_float(value)
            try:
                val = self.format_parser(value, sheet)
                self.set_temp_data(old_value, val, sheet)
                self.set_date(old_value, val, sheet)
                return val
            except:
                # print(sheet + "\t" + old_value + "\t" + str(value) + "\t")
                pass
            if old_value == "输出":
                old_value = "A1"
            reference = ""
            source_r, source_c = ret_row_column(old_value)
            if sheet != "财报" or (sheet == "财报" and (source_c < 8 or (source_c >= 8 and count_func != 0))):
                # 解析函数
                func = formulas.Parser().ast(value)[1].compile()
                func_inputs = func.inputs
                #  遇见偏移函数，解析偏移标示位
                if "OFFSET" in func.outputs[0]:
                    str_o = func.outputs[0]
                    sps = str_o.split("OFFSET(")[1]
                    reference = sps.split(",")[0]
            else:
                count_func += 1
                # 读取解析好的函数
                func = self.func_dic[source_r]
                # 判断是否是最后一个值
                jianshe_time = self.get_temp_data("E230", "假设表")
                end_time = jianshe_time + 240 + 12 + 7
                source_c_char = row_list[source_c]
                source_c_char_1 = row_list[source_c - 1]
                source_c_char__1 = row_list[source_c + 1]
                func_inputs = []
                if source_c == end_time and (source_r in [228, 254, 260, 265, 270, 275, 280, 287, 292, 297, 302]):
                    func = self.func_dic["if_func"]
                    for i in func.inputs:
                        if i == "A1":
                            i = source_c_char + str(source_r - 2)
                        if i == "A2":
                            i = source_c_char + str(source_r - 1)
                        if i == "H18":
                            i = source_c_char + "18"
                        func_inputs.append(i)
                elif source_r == 530 and source_c > 8:
                    func = self.func_dic["I530"]
                    func_inputs.append(source_c_char_1 + "530")
                else:
                    func_outputs = func.outputs[0]
                    if source_r == 217:
                        for i in func.inputs:
                            if i == "H217":
                                func_inputs.append("H217:" + source_c_char_1 + "217")
                                continue
                            new_input = list(i)
                            for i, v in enumerate(i):
                                if v == "I":
                                    new_input[i] = source_c_char
                            func_inputs.append("".join(new_input))
                        if "OFFSET" in func_outputs:
                            str_o = func_outputs
                            sps = str_o.split("OFFSET(")[1]
                            reference = list(sps.split(",")[0])
                            for i, v in enumerate(reference):
                                if v == "I":
                                    reference[i] = source_c_char
                            reference = "".join(reference)
                    else:
                        for i in func.inputs:
                            sums = self.ret_sum(source_r, i, source_c_char, source_c_char_1, source_c_char__1)
                            if sums:
                                func_inputs.append(sums)
                                continue

                            new_input = list(i)
                            for i, v in enumerate(i):
                                if v == "H":
                                    new_input[i] = source_c_char
                                if v == "G":
                                    new_input[i] = source_c_char_1
                                if v == "I":
                                    new_input[i] = source_c_char__1
                            func_inputs.append("".join(new_input))
                        if "OFFSET" in func_outputs:
                            str_o = func_outputs
                            sps = str_o.split("OFFSET(")[1]
                            reference = list(sps.split(",")[0])
                            for i, v in enumerate(reference):
                                if v == "H":
                                    reference[i] = source_c_char
                                if v == "G":
                                    reference[i] = source_c_char_1
                                if v == "I":
                                    reference[i] = source_c_char__1
                            reference = "".join(reference)

            inputs = []
            for input in func_inputs:
                res = self.input_func(sheet, reference, input, old_value)
                inputs.append(res)

            val = self.eng(func(*inputs), sheet, old_value)
            if isinstance(val, str) and val.startswith("="):
                val = self.eng(val, sheet, old_value, count_func=count_func)
            return val
        else:
            # print("写入")
            # 假设表 其他机组类型为空时设置0
            if not value and sheet != "假设表" or (sheet == "假设表" and (old_value == "E19" or old_value == "E20")):
                value = 0
            if str(value) == "#VALUE!" or str(value) == "#REF!" or str(value) == "0" or str(value) == "[[#VALUE!]]" \
                    or str(value) == "#DIV/0!":
                value = 0
            if (old_value == "H195" or old_value == "E247" or old_value == "E246") \
                    and sheet == "假设表" and value == "":
                value = 0
            if (old_value == "E41" or old_value == "E247" or old_value == "E246") \
                    and sheet == "假设表" and value == "":
                value = 0

            # print(sheet + "\t" + old_value + "\t" + str(value) + "\t")
            value = check_value(type(value))(value)
            self.set_temp_data(old_value, value, sheet)
            self.set_date(old_value, value, sheet)
            if "A1" == old_value:
                self.save_all_data()
            # if sum_count == (2 + save_count) * 10000:
            #     save_count += 2
            #     save_all_data()
            return value

    def ret_sum(self, source_r, i, source_c_char, source_c_char_1, source_c_char__1):
        # 当前行有存在异常操作
        sum_dict = {31: {"G31": "G31:" + source_c_char_1 + "31"},
                    32: {"G32": "G32:" + source_c_char_1 + "32"},
                    19: {"H16": "H16:" + source_c_char + "16"},
                    75: {"H74": "H74:" + source_c_char + "74"},
                    93: {"H92": "H92:" + source_c_char + "92"},
                    396: {"G396": "G396:" + source_c_char_1 + "396"},
                    9: {"G8:H8": "G8:" + source_c_char + "8"},
                    77: {"G76:H76": "G76:" + source_c_char + "76"},
                    95: {"G94:H94": "G94:" + source_c_char + "94"},
                    84: {"G76:H76": "G76:" + source_c_char + "76"},
                    102: {"G94:H94": "G94:" + source_c_char + "94"},
                    }
        try:
            return sum_dict.setdefault(source_r, None).setdefault(i, None)
        except:
            return None

    def save_all_data(self):
        # 输出到excel
        with open("data_all_json.txt", "w", encoding="utf-8") as f:
            f.write(json.dumps(self.temp_data))
        with open("data_json.txt", "w", encoding="utf-8") as f:
            f.write(str(self.data))

    def get_return(self):

        dicts = {"概算信息": {}}
        for i in range(5, 21):
            dicts["T" + str(i)] = self.eng("=T" + str(i), "参数", "输出")
        # dicts={'T5': 4.8, 'T6': 1802.5505, 'T7': 26169.7743, 'T8': 5452.0363, 'T9': 5558.3751511095625, 'T10': 2791.6667,
        # 'T11': 546.1333, 'T12': 94.7253, 'T13': 0.57, 'T14': 0.359211072706416, 'T15': 0.3413503971808932,
        # 'T16': 8.133733333333334, 'T17': 38777.6096, 'T18': 31630.6421, 'T19': 0.1174, 'T20': 0.2854}

        self.return_data["项目容量"] = dicts['T5']
        self.return_data["等效小时数"] = dicts['T6']
        self.return_data["静态总投资"] = dicts['T7']
        self.return_data["单位静态总投资"] = dicts['T8']
        self.return_data["单位动态总投资"] = dicts['T9']
        self.return_data["风机价格"] = dicts['T10']
        self.return_data["塔筒价格"] = dicts['T11']
        self.return_data["基础价格"] = dicts['T12']
        self.return_data["电价（含税）"] = dicts['T13']
        self.return_data["度电成本(VAT&CIT)"] = dicts['T14']
        self.return_data["LCOE平准化度电成本"] = dicts['T15']
        self.return_data["Pt（投资回收期）"] = dicts['T16']
        self.return_data["全投资净现值@8.0%"] = dicts['T17']
        self.return_data["资本金净现值@10.0%"] = dicts['T18']
        self.return_data["PIRR（全投资内部收益率）"] = dicts['T19']
        self.return_data["EIRR（资本金内部收益率）"] = dicts['T20']
        self.return_data["概算信息"] = {}
        for i in range(4, 35):
            a = self.get_temp_data("J" + str(i), "概算表")
            if a == -999999:
                a = self.eng("=J" + str(i), "概算表", "输出")
            dicts["概算信息"]["J" + str(i)] = a
            self.return_data["概算信息"]["J" + str(i)] = a
        print(dicts)
        with open(sys.path[0] + "/return_res.txt", "w", encoding="utf-8") as f:
            f.write(json.dumps(self.return_data))

        return dicts

    def get_all_caibao_func(self):
        for i in range(1, 691):
            value = self.get_date_value('H' + str(i), "财报")
            if i == 217:
                value = self.get_date_value('I' + str(i), "财报")
            if value:
                try:
                    func = formulas.Parser().ast(value)[1].compile()
                    self.func_dic[i] = func
                except:
                    if "%" in value:
                        value = percentile_tu_float(value)
                    func = formulas.Parser().ast(value)[1].compile()
                    self.func_dic[i] = func
        value = "=IF(H$18=$F$8,A1+A2,)"
        func = formulas.Parser().ast(value)[1].compile()
        self.func_dic["if_func"] = func
        value = "=H530+1"
        func = formulas.Parser().ast(value)[1].compile()
        self.func_dic["I530"] = func


def get_date_value(data, value: str, sheet: str):
    # 获取数据
    return data[sheet].setdefault(value, None)


import line_profiler

if __name__ == '__main__':
    a = {"power_limiting_year": 20.0, "pit_road": 0, "power_limiting_rate": 0.07, "ten_switching": 0.0,
         "terrain": "\u5e73\u539f", "deadline": 15.0, "anchor_price": 12000.0, "total_concreteprice": 90.0,
         "foundation_concrete1": 64.3, "unit_price": [3650], "market_power_ratio": 0.0,
         "mode_repayment": "\u7b49\u989d\u672c\u91d1", "province": "\u5c71\u4e1c", "setline_len": 0,
         "construction_time": 12, "financing_way": "\u94f6\u884c\u957f\u671f\u501f\u6b3e", "company": "\u53ef\u7814",
         "market_year": 20.0, "market_power_loss": 0.2, "long_term_financingrate": 0.049, "send_project": 0.0,
         "pool": ["GW140/2500-90A"], "nbooster": "110kV\u5347\u538b\u7ad9", "base_reinforcement": 747.4,
         "cut_hour": 2584.4526638297884, "tower_price": 9500.0, "fannum": [47], "total_captical": 117.5,
         "thirtyfive_switching": 0.0}
    p = line_profiler.LineProfiler(ExcelMain.eng)
    p.enable()
    main = ExcelMain(a)
    main.get_return()
    p.disable()
    p.print_stats()

    # main.format_parser("aaa>bbb",1)
