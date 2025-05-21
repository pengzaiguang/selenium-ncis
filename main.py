import os
import time
import random
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.ui import Select
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
import urllib.request
import ddddocr
from enum import Enum
from io import StringIO
from contextlib import redirect_stdout
from datetime import datetime, timedelta
import traceback

# 
js = """
    var date = document.getElementById(arguments[0]);
    date.readOnly = false;
"""

def convertDateTime(date, time = None):
    if (time):
        return "%s-%s-%s %s" % (date[0:4], date[5:7], date[8:10], time)
    else:
        return "%s-%s-%s %s" % (date[0:4], date[5:7], date[8:10], date[11:19])

def correctFee(fee):
    if float(fee) < 0:
        return 0.0
    else:
        return fee
    
def clickCheckboxes(driver, id):
    spans = driver.find_elements(By.XPATH, "//input[@id='%s']/following-sibling::span" % id)
    spans = spans[:-1]
    for span in spans:
        span.click()

def execute():
    driver = None  # 初始化 driver 为
    xls_cnt = 0
    xls_file_name = None  # 初始化
    #for file_name in os.listdir(os.path.join(os.getcwd(), os.pardir)):
    for file_name in os.listdir(os.getcwd()):
        if file_name.startswith("~$"):
            continue
        if file_name.endswith("xls") or file_name.endswith("xlsx"):
            xls_file_name = file_name
            xls_cnt += 1
    pass
    if xls_cnt > 1:
        print("确保当前目录下只有一个EXCEL表格")
        if driver:
            driver.close()
        exit()
    pass
    if not xls_file_name:
        print("没有在当前目录下找到EXCEL表格")
        if driver:
            driver.close()
        exit()
    pass

    #init browser
    options = Options()
    options.add_argument("--disable-extensions")
    options.optionsbinary_location = os.path.join(os.getcwd(), "chrome-win64\\chrome.exe")
    service = Service(os.path.join(os.getcwd(), "chromedriver-win64\\chromedriver.exe"))
    driver = webdriver.Chrome(service = service, options = options)
    driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source": """
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            })
        """
    })
    driver.maximize_window()
    driver.get("https://quality.ncis.cn/report-disease/drgs")
    time.sleep(random.uniform(0.68, 1.28))
    enter = driver.find_element(By.XPATH, '//span[text()="国家单病种质量管理与控制平台"]')
    if enter:   # 需要登录
        enter.click()
        auto_login(driver)
    pass

    df = pd.read_excel(os.path.join(os.getcwd(), xls_file_name), dtype=str)
    for index, row in df.iterrows():
        #print(row)
        if len(str(row['住院号']))  < 5:
            continue

        # 进入DRGS页面
        driver.get("https://quality.ncis.cn/report-disease/drgs")
        time.sleep(random.uniform(0.68, 1.38))
        try:
            driver.find_element(By.XPATH, '//label[contains(text(), "不再提示")]/span/input').click()
        except:
            pass
        try:
            driver.find_element(By.CLASS_NAME, 'ivu-icon.ivu-icon-ios-close').click()
        except:
            pass

        # 选择类别和病种
        category = row['系统'].replace("或", "/")
        category = category.replace("（", "(")
        category = category.replace("）", ")")
        try:
            #first_element = driver.find_element(By.XPATH, f"//li[contains(text(),‘{category}’)]")
            WebDriverWait(driver, 30).until(EC.visibility_of_element_located((By.XPATH, f"//li[contains(text(),'{category}')]"))).click()
        except:
            print("*** 暂不支持自动录入的疾病类型：%s" % category)
            continue
        pass
        try:
            driver.find_elements(By.XPATH, f"//div[contains(text(),'{row['病种']}')]/parent::div/div")[1].click()
        except:
            print("*** 暂不支持自动录入的疾病类型：%s" % row['病种'])
            continue
        pass

        # 进入录入页面
        element = WebDriverWait(driver, 60).until(EC.presence_of_element_located((By.ID, "myiframe")))
        time.sleep(random.uniform(1.0, 1.68))
        driver.switch_to.frame(element)
        WebDriverWait(driver, 300).until(EC.presence_of_element_located((By.ID, "submit")))

        # 填写基本信息
        fillSuccess = True
        try:
            # 病种分发
            match row['病种']:
                case '围手术期预防深静脉血栓栓塞':
                    disease_perioperative_prevention_of_deep_vein_thrombosis(driver, row)
                case '围手术期预防感染':
                    disease_perioperative_infection_prophylaxis(driver, row)
                case '异位妊娠':
                    disease_ectopic_pregnancy(driver, row)
                case '子宫肌瘤':
                    disease_uterine_fibroids(driver, row)
                case '宫颈癌（手术治疗）':
                    disease_cervical_cancer(driver, row)
                case _:
                    print("*** 未知疾病种类")
            pass
        except Exception as e:
            fillSuccess = False
            print("Exception: ", e)
            traceback.print_exc()

        if fillSuccess:
            # submit
            # submit = driver.find_element(By.ID, "submit")
            # driver.execute_script("arguments[0].focus()", submit)
            # submit.click()
            # time.sleep(random.uniform(0.68, 1.38))
            # confirm = driver.find_elements(By.CLASS_NAME, "layui-layer-btn1")
            # if len(confirm) > 0:
            #     confirm[0].click()
            #     try:
            #         WebDriverWait(driver, 5).until(EC.invisibility_of_element(submit))
            #         print(" ### 成功提交")
            #     except Exception as e:
            #         print(" *** 数据重复，已跳过")
            # else:
            #     input(" *** 自动填充失败，请自行填充数据并提交，并在本命令行窗口按回车键继续。\n *** 如需跳过该条数据，直接按回车。")
            # pass

            # save
            save = driver.find_element(By.ID, "save")
            driver.execute_script("arguments[0].focus()", save)
            save.click()
            time.sleep(random.uniform(0.68, 1.38))
        pass
    pass    # end for row

    # 关闭浏览器
    input(" *** 运行结束，按回车键退出。")
    # driver.close()

# 自动登录
def auto_login(driver):
    with open(os.path.join(os.getcwd(), "password.txt"), 'r', encoding='utf-8') as f:
        lines = f.readlines()
    pass
    username = lines[0].strip()
    password = lines[1].strip()
    username_input = driver.find_element(By.XPATH, '//input[@placeholder="账户"]')
    password_input = driver.find_element(By.XPATH, '//input[@type="password"]')
    username_input.send_keys(username)
    password_input.send_keys(password)

    graph_element = WebDriverWait(driver, 20).until(EC.visibility_of_element_located((By.CLASS_NAME, "code-img")))
    security_input = driver.find_element(By.XPATH, '//input[@placeholder="验证码"]')
    img_dir = os.path.join(os.getcwd(), "img")
    if not os.path.isdir(img_dir):
        os.makedirs(img_dir)
    graph_url = graph_element.get_attribute("src")
    graph_path = os.path.join(img_dir, "%s.gif" % datetime.now().strftime("%Y-%m-%d-%H-%M-%S"))
    urllib.request.urlretrieve(graph_url, graph_path)
    with open(graph_path, 'rb') as f:
        img_bytes = f.read()
    pass
    output_string = StringIO()
    with redirect_stdout(output_string):
        ocr = ddddocr.DdddOcr()
        security_code = ocr.classification(img_bytes)
        security_input.send_keys(security_code)
    pass

    login = driver.find_element(By.XPATH, '//span[text()="登录"]/parent::button')
    login.click()

    try:
        WebDriverWait(driver, 60).until(EC.invisibility_of_element(login))
        # wait = WebDriverWait(driver, timeout = 60, poll_frequency = 1)
        # wait.until(EC.invisibility_of_element_located((By.XPATH, '//span[text()="登录"]/parent::button')))
    except Exception as e:
        print(f"Exception 的值为: {e}")
    pass

# 基本信息
def basic_info(driver, row):
    # 质控医师
    quality_ctrl_doctor_input = driver.find_element(By.ID, "create_CM_1")
    quality_ctrl_doctor_input.send_keys(row['质控医师'].replace(" ", ""))

    # 质控护士
    quality_ctrl_nurse_input = driver.find_element(By.ID, "create_CM_2")
    quality_ctrl_nurse_input.send_keys(row['质控护士'].replace(" ", ""))

    # 主治医师
    attending_doctor_input = driver.find_element(By.ID, "create_CM_3")
    attending_doctor_input.send_keys(row['主治医师'].replace(" ", ""))

    # 责任护士
    primary_nurse_input = driver.find_element(By.ID, "create_CM_4")
    primary_nurse_input.send_keys(row['责任护士'].replace(" ", ""))

    # 上报科室
    try:
        submit_department_input = driver.find_element(By.ID, "create_CM_186")
        submit_department_input.send_keys(row['出院科室'])
    except Exception as e:
        pass
    try:
        submit_department_input = driver.find_element(By.ID, "create_DVT_223")
        submit_department_input.send_keys(row['出院科室'])
    except Exception as e:
        pass

    # 患者病案号
    patient_case_number_input = driver.find_element(By.ID, "create_CM_5")
    patient_case_number_input.send_keys(row['住院号'])

    # 患者身份证号
    patient_id_number_input = driver.find_element(By.ID, "create_CM_6")
    patient_id_number_input.send_keys(row['身份证'])

    # 主要诊断ICD-10四位亚目编码与名称
    try:
        primary_diagnosis_index_4_select = driver.find_element(By.ID, "create_CM_7")
        primary_diagnosis_index_4_select.send_keys(row['主诊ICD码'][0:5])
    except Exception as e:
        pass

    # 主要诊断ICD-10六位临床扩展编码与名称
    try:
        primary_diagnosis_index_6_select = driver.find_element(By.ID, "create_CM_8")
        primary_diagnosis_index_6_select.send_keys(row['主诊ICD码'])
    except Exception as e:
        pass    

    # 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称
    primary_surgery_index_4_select = driver.find_element(By.ID, "create_CM_9")
    primary_surgery_index_4_select.send_keys(row['第一个手术码'][0:5])
    
    # 主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称
    primary_surgery_index_6_select = driver.find_element(By.ID, "create_CM_10")
    primary_surgery_index_6_select.send_keys(row['第一个手术码'])

    # 是否出院后31天内重复住院
    repeat_hospitalization_input = driver.find_elements(By.XPATH, '//input[@id="create_CM_11"]/parent::div/span')[2]
    repeat_hospitalization_input.click()
    # driver.execute_script("arguments[0].scrollIntoView();", repeat_hospitalization_input)
    # time.sleep(random.uniform(0.68, 1.38))

    # 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称
    # primary_surgery_index_4_select = driver.find_element(By.ID, "create_CM_9")
    #primary_surgery_index_4_select.send_keys(row['第一个手术码'][0:5])
    # Select(primary_surgery_index_4_select).select_by_index(1)

    # 出生日期时间
    # js = """
    #     var date = document.getElementById(arguments[0]);
    #     date.readOnly = false;
    #     """
    # driver.execute_script(js, "create_CM_13")
    # birthday_date_input = driver.find_element(By.ID, "create_CM_13")
    # birthday_date_input.send_keys(row['出生日期'].split(" ")[0])
    # print(row['出生日期'])

    isMale = '男' == row['性别']

    # 患者性别
    try:
        gender = "F" if isMale else "F"
        driver.find_element(By.XPATH, f"//span[@key='{gender}']").click()
    except Exception as e:
        pass

    # 患者体重
    if isMale:
        weight = 65 + random.randint(-5, 15)
    else:
        weight = 50 + random.randint(-5, 10)
    patient_weight_input = driver.find_element(By.ID, "create_CM_15")
    patient_weight_input.send_keys(weight)

    # 患者身高
    if isMale:
        if weight >= 75:
            height = 170 + random.randint(0, 15)
        else:
            height = 170 + random.randint(-5, 0)
    else:
        if weight >= 55:
            height = 150 + random.randint(10, 15)
        else:
            height = 150 + random.randint(0, 10)
    patient_height_input = driver.find_element(By.ID, "create_CM_227")
    patient_height_input.send_keys(height)

    # 到达本院急诊或者门诊日期时间是否无法确定或无记录
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "到达本院急诊或者门诊日期时间是否无法确定或无记录")]/following-sibling::span[@key="UTD"]').click()
    except Exception as e:
        pass

    # 入院日期时
    driver.execute_script(js, "create_CM_16")
    random_hour = random.randint(8, 22)
    admission_time = f"{random_hour:02d}:00"
    admission_date = convertDateTime(row['入院日期'], admission_time)
    admission_date_input = driver.find_element(By.ID, "create_CM_16")
    row['admission_date'] = datetime.strptime(admission_date, "%Y-%m-%d %H:%M")
    admission_date_input.send_keys(admission_date)

    # 出院日期时间
    driver.execute_script(js, "create_CM_17")
    random_hour = random.randint(9, 12)
    discharge_time = f"{random_hour:02d}:00"
    discharge_date = convertDateTime(row['出院日期'], discharge_time)
    discharge_date_input = driver.find_element(By.ID, "create_CM_17")
    discharge_date_input.send_keys(discharge_date)

    # 手术开始时间
    driver.execute_script(js, "create_CM_24")
    surgery_begin_time_input = driver.find_element(By.ID, "create_CM_24")
    surgery_begin_time_obj = datetime.strptime(admission_date, "%Y-%m-%d %H:%M")
    delta_day = 0 #random.randint(1,2)
    delta_hour = random.randint(8, 12)
    delta_minute = random.randint(0, 59)
    delta = timedelta(days = delta_day, hours = delta_hour, minutes = delta_minute)
    surgery_begin_time_obj += delta
    row['surgery_begin_time_obj'] = surgery_begin_time_obj
    surgery_begin_time_input.send_keys(surgery_begin_time_obj.strftime('%Y-%m-%d %H:%M'))

    # 手术结束时间
    driver.execute_script(js, "create_CM_25")
    surgery_end_time_input = driver.find_element(By.ID, "create_CM_25")
    delta_hour = 1
    delta_minute = random.randint(10, 60)
    delta = timedelta(hours = delta_hour, minutes = delta_minute)
    surgery_end_time_obj = surgery_begin_time_obj + delta
    row['surgery_end_time_obj'] = surgery_end_time_obj
    surgery_end_time_input.send_keys(surgery_end_time_obj.strftime('%Y-%m-%d %H:%M'))

    # 费用支付方式
    payment_method_select = driver.find_element(By.ID, "create_CM_28")
    payment_method_select.send_keys(row['付款方式'].replace("城乡", "城镇"))
    # 收入住院途径
    admission_route_select = driver.find_element(By.ID, "create_CM_29")
    admission_route_select.send_keys('门诊')
    # 到院交通工具
    transportation_method_select = driver.find_element(By.ID, "create_CM_30")
    transportation_method_select.send_keys(random.choice(["私家车", "出租车", "其它"]))
    # 离院方式选择
    leave_hospital_method_input = driver.find_element(By.ID, "create_CM_79")
    leave_hospital_method_input.send_keys("医嘱离院")
    

#  基本费用
def basic_fee(driver, row):
    # 住院总费用
    total_cost_input = driver.find_element(By.ID, "create_CM_98")
    driver.execute_script("arguments[0].scrollIntoView();", total_cost_input)
    time.sleep(random.uniform(0.3, 0.6))  # 等待页面滚动
    total_cost_input.send_keys(correctFee(row['总费用']))

    # 住院总费用中自付金额
    self_cost_input = driver.find_element(By.ID, "create_CM_99")
    self_fee = row['自付金额']
    if float(row['自付金额']) > float(row['总费用']):
        self_fee = row['总费用']
    self_cost_input.send_keys(correctFee(self_fee))

    # 一般医疗服务费
    general_medical_service_fee_input = driver.find_element(By.ID, "create_CM_100")
    general_medical_service_fee_input.send_keys(correctFee(row['一般医疗服务费']))

    # 一般治疗操作费
    general_treatment_operation_fee_input = driver.find_element(By.ID, "create_CM_101")
    general_treatment_operation_fee_input.send_keys(correctFee(row['一般治疗操作费']))

    # 护理费
    nursing_fee_input = driver.find_element(By.ID, "create_CM_102")
    nursing_fee_input.send_keys(correctFee(row['护理费']))

    # 综合医疗服务类其他费用
    comprehensive_medical_service_other_fee_input = driver.find_element(By.ID, "create_CM_103")
    comprehensive_medical_service_other_fee_input.send_keys(row['其他费用'])

    # 病理诊断费
    pathology_diagnosis_fee_input = driver.find_element(By.ID, "create_CM_104")
    pathology_diagnosis_fee_input.send_keys(correctFee(row['病理诊断费']))

    # 实验室诊断费
    laboratory_diagnosis_fee_input = driver.find_element(By.ID, "create_CM_105")
    laboratory_diagnosis_fee_input.send_keys(correctFee(row['实验室诊断费']))

    # 影像学诊断费
    diagnosis_imaging_fee_input = driver.find_element(By.ID, "create_CM_106")
    diagnosis_imaging_fee_input.send_keys(correctFee(row['影像学诊断费']))

    # 临床诊断项目费
    clinical_diagnosis_program_fee_input = driver.find_element(By.ID, "create_CM_107")
    clinical_diagnosis_program_fee_input.send_keys(correctFee(row['临床诊断项目费']))

    # 非手术治疗项目费
    non_surgical_treatment_program_fee_input = driver.find_element(By.ID, "create_CM_108")
    non_surgical_treatment_program_fee_input.send_keys(correctFee(row['非手术治疗项目费']))

    # 临床物理治疗费
    clinical_physiotherapy_fee_input = driver.find_element(By.ID, "create_CM_109")
    clinical_physiotherapy_fee_input.send_keys(correctFee(row['非手术治疗项目费其中临床物理治疗费']))

    # 手术治疗费
    surgical_treatment_fee_input = driver.find_element(By.ID, "create_CM_110")
    surgical_treatment_fee_input.send_keys(correctFee(row['手术治疗费']))

    # 麻醉费
    anesthesia_fee_input = driver.find_element(By.ID, "create_CM_111")
    anesthesia_fee_input.send_keys(correctFee(row['手术治疗费其中麻醉费']))

    # 手术费
    surgery_fee_input = driver.find_element(By.ID, "create_CM_112")
    surgery_fee_input.send_keys(correctFee(row['手术治疗费其中手术费']))

    # 康复费
    rehabilitation_fee_input = driver.find_element(By.ID, "create_CM_113")
    rehabilitation_fee_input.send_keys(correctFee(row['康复费']))

    # 中医治疗费
    tcm_fee_input = driver.find_element(By.ID, "create_CM_114")
    tcm_fee_input.send_keys(correctFee(row['中医治疗类']))

    # 西药费
    western_medicine_fee_input = driver.find_element(By.ID, "create_CM_115")
    western_medicine_fee_input.send_keys(correctFee(row['西药费']))

    # 抗菌药物费
    antibacterial_drug_fee_input = driver.find_element(By.ID, "create_CM_116")
    antibacterial_drug_fee_input.send_keys(correctFee(row['西药费其中抗菌药物费用']))

    # 中成药费
    proprietary_chinese_medicine_fee_input = driver.find_element(By.ID, "create_CM_117")
    proprietary_chinese_medicine_fee_input.send_keys(correctFee(row['中成药费']))

    # 中草药费
    chinese_herbal_medicine_fee_input = driver.find_element(By.ID, "create_CM_118")
    chinese_herbal_medicine_fee_input.send_keys(correctFee(row['中草药费']))

    # 血费
    blood_cost_input = driver.find_element(By.ID, "create_CM_119")
    blood_cost_input.send_keys(correctFee(row['血费']))

    # 白蛋白类制品费
    albumin_product_fee_input = driver.find_element(By.ID, "create_CM_120")
    albumin_product_fee_input.send_keys(correctFee(row['白蛋白类制品费']))

    # 球蛋白类制品费
    globulin_product_input = driver.find_element(By.ID, "create_CM_121")
    globulin_product_input.send_keys(correctFee(row['球蛋白制品费']))

    # 凝血因子类制品费
    coagulation_factor_product_fee_input = driver.find_element(By.ID, "create_CM_122")
    coagulation_factor_product_fee_input.send_keys(correctFee(row['凝血因子类制品费']))

    # 细胞因子类制品费
    cytokine_product_fee_input = driver.find_element(By.ID, "create_CM_123")
    cytokine_product_fee_input.send_keys(correctFee(row['细胞因子类费']))

    # 检查用一次性医用材料费
    disposable_medical_metrials_for_examination_fee_input = driver.find_element(By.ID, "create_CM_124")
    disposable_medical_metrials_for_examination_fee_input.send_keys(correctFee(row['检查用一次性医用材料费']))

    # 治疗用一次性医用材料费
    disposable_medical_metrials_for_treatment_fee_input = driver.find_element(By.ID, "create_CM_125")
    disposable_medical_metrials_for_treatment_fee_input.send_keys(correctFee(row['治疗用一次性医用材料费']))

    # 手术用一次性医用材料费
    disposable_medical_metrials_for_surgery_fee_input = driver.find_element(By.ID, "create_CM_126")
    disposable_medical_metrials_for_surgery_fee_input.send_keys(correctFee(row['手术用一次性医用材料费']))

    # 其他费
    other_fee_input = driver.find_element(By.ID, "create_CM_127")
    other_fee_input.send_keys(correctFee(row['其他费']))
    pass


# 住院期间为患者提供术前、术后健康教育与出院时提供教育告知五要素情况
def check_pre_post_op_health_education(driver, row):
    # 术前：健康辅导
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "术前：健康辅导")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 术后：健康辅导
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "术后：健康辅导")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 交与患者“出院小结”的副本告知患者出院时风险因素
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "交与患者“出院小结”的副本告知患者出院时风险因素")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 出院带药
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "出院带药")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
     # 告知何为发生紧急意外情况或者疾病复发
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "告知何为发生紧急意外情况或者疾病复发")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 告知发生紧急情况时求援救治途径
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "告知发生紧急情况时求援救治途径")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "告知发生紧急意外情况或者疾病复发如何救治及前途经")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 出院时教育与随访
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "出院时教育与随访")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    # 告知何为风险因素与紧急情况
    try:
        driver.find_element(By.XPATH, '//span[contains(text(), "告知何为风险因素与紧急情况")]/following-sibling::span[@key="a"]').click()
    except Exception as e:
        pass
    pass

# 手术切口愈合情况
def check_surgical_wound_healing(driver, row):
    # 手术野皮肤准备常用方法的选择
    create_CM_72_input = driver.find_element(By.ID, "create_CM_72")
    create_CM_72_input.send_keys('剪刀清除毛发')
    # 使用含抗菌剂（三氯生）缝线
    create_CM_73_input = driver.find_element(By.ID, "create_CM_73")
    create_CM_73_input.send_keys('抗菌薇乔®')
    # 手术切口类别的选择
    create_CM_74_input = driver.find_element(By.ID, "create_CM_74")
    create_CM_74_input.send_keys('Ⅱ类切口')
    # 手术切口愈合情况的选择
    create_CM_75_input = driver.find_element(By.ID, "create_CM_75")
    create_CM_75_input.send_keys('甲级愈合')
    pass

# 离院方式
def check_discharge_method(driver, row):
    # 离院方式选择
    create_CM_79_input = driver.find_element(By.ID, "create_CM_79")
    create_CM_79_input.send_keys('医嘱离院')
    pass

# 患者对服务的体验与评价
def patient_service_evaluation(driver, row):
    # 患者是否对服务的体验与评价
    # driver.find_elements(By.XPATH, "//input[@id='create_CM_85']/parent::div/span")[2].click()
    driver.find_element(By.XPATH, '//span[contains(text(), "患者是否对服务的体验与评价")]/following-sibling::span[@key="n"]').click()
    pass


# 围手术期预防深静脉血栓栓塞
def disease_perioperative_prevention_of_deep_vein_thrombosis(driver, row):
    # 基本信息
    basic_info(driver, row)
    # 手术时间
    surgery_begin_time_obj = row['surgery_begin_time_obj']
    surgery_end_time_obj = row['surgery_end_time_obj']

    # DVT-1 预防性抗菌药物使用情况
    # 是否使用预防性抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "是否使用预防性抗菌药物")]/following-sibling::span[@key="y"]').click()
    # 预防性抗菌药物选择
    driver.find_element(By.XPATH, '//span[contains(text(), "第一代或第二代头孢菌素")]').click()
    # 使用首剂抗菌药物起始时间
    driver.execute_script(js, "create_CM_41")
    delta = timedelta(minutes = 30)
    first_dose_of_antimicrobial = surgery_begin_time_obj - delta
    create_CM_41_input = driver.find_element(By.ID, "create_CM_41")
    create_CM_41_input.send_keys(first_dose_of_antimicrobial.strftime('%Y-%m-%d %H:%M'))
    # 术中追加抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "手术时间是否≥3小时")]/following-sibling::span[@key="n"]').click()
    # driver.find_element(By.XPATH, '//span[contains(text(), "是否术中追加抗菌药物")]/following-sibling::span[@key="n"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "术中出血量是否≥1500ml")]/following-sibling::span[@key="n"]').click()
    # driver.find_element(By.XPATH, '//span[contains(text(), "术后是否使用抗菌药物")]/following-sibling::span[@key="n"]').click()
    # 术后抗菌药物停止使用时间
    driver.execute_script(js, "create_CM_47")
    delta_hour = random.randint(24, 48)
    delta = timedelta(hours = delta_hour)
    antimicrobial_agents_after_surgery_date = surgery_end_time_obj + delta
    create_CM_47_input = driver.find_element(By.ID, "create_CM_47")
    create_CM_47_input.send_keys(antimicrobial_agents_after_surgery_date.strftime('%Y-%m-%d %H:%M'))

    # DVT-2 预防术后深静脉血栓形成
    # 术前进行Caprini血栓风险因素评估情况
    create_37_input = driver.find_element(By.ID, "create_37")
    create_37_input.send_keys('Caprini血栓风险因素评估')
    # 评估分值
    create_38_input = driver.find_element(By.ID, "create_38")
    create_38_input.send_keys('3')
    # 临床常用检测方法
    driver.find_element(By.XPATH, '//span[contains(text(), "临床常用检测方法")]/following-sibling::span[@key="UTD"]').click()
    # 术后可能诱发危险因素
    driver.find_element(By.XPATH, '//span[contains(text(), "术后可能诱发危险因素")]/following-sibling::span[@key="a"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "术后可能诱发危险因素")]/following-sibling::span[@key="b"]').click()
    # 是否有实施预防术后深静脉血栓措施的禁忌
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有实施预防术后深静脉血栓措施的禁忌")]/following-sibling::span[@key="n"]').click()
    # 是否实施基本预防措施
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施基本预防措施")]/following-sibling::span[@key="y"]').click()
    # 基本预防措施的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "基本预防措施的选择")]/following-sibling::span[@key="a"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "基本预防措施的选择")]/following-sibling::span[@key="b"]').click()
    # 基本预防措施医嘱执行日期
    driver.execute_script(js, "create_DVT_168")
    # delta_hour = random.randint(24, 48)
    delta = timedelta(hours = 8)
    create_DVT_168_date = surgery_begin_time_obj + delta
    create_DVT_168_input = driver.find_element(By.ID, "create_DVT_168")
    create_DVT_168_input.send_keys(create_DVT_168_date.strftime('%Y-%m-%d %H:%M'))
    # 是否实施机械预防措施
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施机械预防措施")]/following-sibling::span[@key="y"]').click()
    # 机械预防措施的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "机械预防措施的选择")]/following-sibling::span[@key="b"]').click()
    # 机械预防措施医嘱执行起始日期
    driver.execute_script(js, "create_DVT_172")
    # delta_hour = random.randint(24, 48)
    delta = timedelta(hours = 8)
    create_DVT_172_date = surgery_begin_time_obj + delta
    create_DVT_172_input = driver.find_element(By.ID, "create_DVT_172")
    create_DVT_172_input.send_keys(create_DVT_172_date.strftime('%Y-%m-%d %H:%M'))
    # 是否实施药物预防措施
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施药物预防措施")]/following-sibling::span[@key="y"]').click()
    # 预防性地药物的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "预防性地药物的选择")]/following-sibling::span[@key="c"]').click()
    # 预防性地药物医嘱执行日期
    driver.execute_script(js, "create_DVT_173")
    # delta_hour = random.randint(24, 48)
    delta = timedelta(hours = 24)
    create_DVT_173_date = surgery_begin_time_obj + delta
    create_DVT_173_input = driver.find_element(By.ID, "create_DVT_173")
    create_DVT_173_input.send_keys(create_DVT_173_date.strftime('%Y-%m-%d %H:%M'))
    # 出院后继续使用抗凝药
    create_DVT_174_input = driver.find_element(By.ID, "create_DVT_174")
    create_DVT_174_input.send_keys('无法确定或无记录')
    # 术前、术后、出院时为患者提供针对性健康教育服务
    driver.find_element(By.XPATH, '//span[contains(text(), "术前、术后、出院时为患者提供针对性健康教育服务")]/following-sibling::span[@key="a"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "术前、术后、出院时为患者提供针对性健康教育服务")]/following-sibling::span[@key="b"]').click()
    
    # DVT-3 手术后并发症
    # 是否有手术后并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术后并发症")]/following-sibling::span[@key="n"]').click()

    # DVT-4 手术切口愈合情况
    check_surgical_wound_healing(driver, row)

    # DVT-5 离院方式
    # 离院方式选择
    check_discharge_method(driver, row)

    # DVT-6 患者对服务的体验与评价
    # 患者是否对服务的体验与评价
    patient_service_evaluation(driver, row)

    # DVT-7 住院费用
    basic_fee(driver, row)
    pass

# 围手术期预防感染
def disease_perioperative_infection_prophylaxis(driver, row):
    # 手术类型
    surgery_type_selector = driver.find_element(By.ID, "create_PIP_230")
    surgery_type_selector.send_keys('住院手术')
    # 基本信息
    basic_info(driver, row)
    # 主要手术操作栏中提取ICD-9-CM-3四位亚目编码与名称(心血管病医院)
    create_2_select = driver.find_element(By.ID, "create_2")
    create_2_select.send_keys(row['第一个手术码'][0:5])
    # 主要手术操作栏中提取ICD-9-CM-3六位临床扩展编码与名称(心血管病医院)：
    create_PIP_229_select = driver.find_element(By.ID, "create_PIP_229")
    create_PIP_229_select.send_keys(row['第一个手术码'])
    # 手术时间
    surgery_begin_time_obj = row['surgery_begin_time_obj']

    # PIP-1: 预防性抗菌药物使用情况
    # 是否使用预防性抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "是否使用预防性抗菌药物")]/following-sibling::span[@key="y"]').click()
    # 预防性抗菌药物选择
    driver.find_element(By.XPATH, '//span[contains(text(), "第一代或第二代头孢菌素")]').click()
    # 使用首剂抗菌药物起始时间
    driver.execute_script(js, "create_CM_41")
    delta = timedelta(minutes = 30)
    create_CM_41_date = surgery_begin_time_obj - delta
    create_CM_41_input = driver.find_element(By.ID, "create_CM_41")
    create_CM_41_input.send_keys(create_CM_41_date.strftime('%Y-%m-%d %H:%M'))
    # 手术时间是否≥3小时
    driver.find_element(By.XPATH, '//span[contains(text(), "手术时间是否≥3小时")]/following-sibling::span[@key="n"]').click()
    # 术中出血量是否≥1500ml
    driver.find_element(By.XPATH, '//span[contains(text(), "术中出血量是否≥1500ml")]/following-sibling::span[@key="n"]').click()
    # 术后是否使用抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "术后是否使用抗菌药物")]/following-sibling::span[@key="n"]').click()

    # PIP-2: 手术后并发症
    # 是否有手术后并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术后并发症")]/following-sibling::span[@key="n"]').click()

    # PIP-3: 手术切口愈合情况
    check_surgical_wound_healing(driver, row)
    
    # PIP-4: 离院方式
    check_discharge_method(driver, row)

    # PIP-5: 患者对服务的体验与评价
    patient_service_evaluation(driver, row)

    # PIP-6: 住院费用
    basic_fee(driver, row)
    pass

# 异位妊娠
def disease_ectopic_pregnancy(driver, row):
    #  基本信息
    basic_info(driver, row)

    admission_date = row['admission_date']
    # 手术时间
    surgery_begin_time_obj = row['surgery_begin_time_obj']
    surgery_end_time_obj = row['surgery_end_time_obj']
    # 末次月经日期是否无法确定或无记录
    driver.execute_script(js, "create_CM_18")
    delta_day = random.randint(40, 60)
    delta = timedelta(days = delta_day)
    last_menstrual_date = surgery_begin_time_obj - delta
    create_CM_18_input = driver.find_element(By.ID, "create_CM_18")
    create_CM_18_input.send_keys(last_menstrual_date.strftime('%Y-%m-%d'))

    # EP-1 患者入院病情评估
    # 高危因素评估评估日期时间
    driver.execute_script(js, "create_30")
    # delta_hour = random.randint(2, 8)
    # delta = timedelta(hours = delta_hour)
    create_30_date = admission_date
    create_30_input = driver.find_element(By.ID, "create_30")
    create_30_input.send_keys(create_30_date.strftime('%Y-%m-%d %H:%M'))
    # 高危因素的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "高危因素的选择")]/following-sibling::span[@key="UTD"]').click()
    # 妊娠周数
    create_35_select = driver.find_element(By.ID, "create_35")
    Select(create_35_select).select_by_index(2)
    # 腹痛程度的选择
    create_36_select = driver.find_element(By.ID, "create_36")
    # Select(create_36_select).select_by_index(2)
    create_36_select.send_keys('隐痛')
    # 生命体征是否平稳
    create_DG_229_select = driver.find_element(By.ID, "create_DG_229")
    create_DG_229_select.send_keys('生命体征平稳')
    # 超声检查途径的选择
    create_39_select = driver.find_element(By.ID, "create_39")
    create_39_select.send_keys('经阴道超声')
    # 超声检查描述
    driver.find_element(By.XPATH, '//span[contains(text(), "超声检查描述：")]/following-sibling::span[@key="UTD"]').click()
    # 子宫内膜厚度(mm)
    create_DG_233_select = driver.find_element(By.ID, "create_DG_233")
    create_DG_233_select.send_keys('10')
    # 输卵管妊娠包块最大径的选择
    create_40_select = driver.find_element(By.ID, "create_40")
    Select(create_40_select).select_by_index(2)
    # 盆腔内出血量最大径的选择
    create_41_select = driver.find_element(By.ID, "create_41")
    Select(create_41_select).select_by_index(2)
    # 到院首次B超检查提示异位妊娠征象报告的时间
    driver.execute_script(js, "create_45")
    # delta_hour = random.randint(4, 8)
    # delta = timedelta(hours = delta_hour)
    create_45_date = admission_date
    create_45_input = driver.find_element(By.ID, "create_45")
    create_45_input.send_keys(create_45_date.strftime('%Y-%m-%d %H:%M'))
    # 是否进行β-HCG测定
    driver.find_element(By.XPATH, '//span[contains(text(), "是否进行β-HCG测定")]/following-sibling::span[@key="n"]').click()
    # 是否进行穿刺
    driver.find_element(By.XPATH, '//span[contains(text(), "是否进行穿刺")]/following-sibling::span[@key="n"]').click()
    # 治疗方式选择
    create_49_select = driver.find_element(By.ID, "create_49")
    create_49_select.send_keys('紧急手术')
    # 需紧急手术的病情严重程度评估
    try:
        create_DG_252_select = driver.find_element(By.ID, "create_DG_252")
        create_DG_252_select.send_keys('无法确定或无记录')
    except:
        pass

    # EP-3 手术治疗情况
    # 手术的指征的选择
    create_70_select = driver.find_element(By.ID, "create_70")
    create_70_select.send_keys('临床病情稳定的患者，或与其他有指征的手术同时进行')
    # 手术方式选择
    create_77_select = driver.find_element(By.ID, "create_77")
    create_77_select.send_keys('腹腔镜手术')
    # 腹腔镜手术术式选择
    create_80_select = driver.find_element(By.ID, "create_80")
    Select(create_80_select).select_by_index(1)

    # EP-4 围术期预防性抗菌药物使用情况
    # 手术治疗的患者，是否预防性用药
    driver.find_element(By.XPATH, '//span[contains(text(), "手术治疗的患者，是否预防性用药")]/following-sibling::span[@key="y"]').click()
    # 预防性抗菌药物选择
    driver.find_element(By.XPATH, '//span[contains(text(), "预防性抗菌药物选择")]/following-sibling::span[@key="d"]').click()
    # 使用首剂抗菌药物起始时间
    driver.execute_script(js, "create_CM_41")
    # delta_minutes = random.randint(20, 40)
    delta = timedelta(minutes = 30)
    create_CM_41_date = surgery_begin_time_obj - delta
    create_CM_41_input = driver.find_element(By.ID, "create_CM_41")
    create_CM_41_input.send_keys(create_CM_41_date.strftime('%Y-%m-%d %H:%M'))
    # 手术时间是否≥3小时
    driver.find_element(By.XPATH, '//span[contains(text(), "手术时间是否≥3小时")]/following-sibling::span[@key="n"]').click()
    # 术中出血量是否≥1500ml
    driver.find_element(By.XPATH, '//span[contains(text(), "术中出血量是否≥1500ml")]/following-sibling::span[@key="n"]').click()
    # 术后抗菌药物停止使用时间
    driver.execute_script(js, "create_CM_47")
    delta_hours = random.randint(12, 24)
    delta = timedelta(hours = delta_hours)
    create_CM_47_date = surgery_end_time_obj + delta
    create_CM_47_input = driver.find_element(By.ID, "create_CM_47")
    create_CM_47_input.send_keys(create_CM_47_date.strftime('%Y-%m-%d %H:%M'))

    # EP-5 术后并发症与再手术情况
    # 是否有手术后并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术后并发症")]/following-sibling::span[@key="n"]').click()

    # EP-6 输血量
    # 术中腹腔内出血量
    delta = random.randint(40, 60) // 10 * 10
    create_120_select = driver.find_element(By.ID, "create_120")
    create_120_select.send_keys(delta)
    # 是否实施输血
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施输血")]/following-sibling::span[@key="n"]').click()

    # EP-7 住院期间为患者提供健康教育与出院时提供教育告知五要素情况
    check_pre_post_op_health_education(driver, row)

    # EP-8 手术切口愈合情况
    check_surgical_wound_healing(driver, row)

    # EP-9 离院方式
    check_discharge_method(driver, row)
    
    # EP-10 患者对服务的体验与评价
    patient_service_evaluation(driver, row)

    # EP-11 住院费用
    basic_fee(driver, row)
    pass

# 子宫肌瘤
def disease_uterine_fibroids(driver, row): 
    #  基本信息
    basic_info(driver, row)
    # 发病日期时间是否无法确定或无记录
    driver.find_element(By.XPATH, '//span[contains(text(), "发病日期时间是否无法确定或无记录")]/following-sibling::span[@key="UTD"]').click()

    # 手术时间
    surgery_begin_time_obj = row['surgery_begin_time_obj']
    surgery_end_time_obj = row['surgery_end_time_obj']

    # UM-1 患者入院病情评估与术式选择
    # 患者评估与知情同意
    create_36_select = driver.find_element(By.ID, "create_36")
    create_36_select.send_keys('有生育要求、期望保留子宫者')
    # 影像学检查评估
    driver.find_element(By.XPATH, '//span[contains(text(), "影像学检查评估")]/following-sibling::span[@key="a"]').click()
    # 肌瘤数目
    create_39_select = driver.find_element(By.ID, "create_39")
    create_39_select.send_keys('多发')
    # 子宫大小
    create_41_select = driver.find_element(By.ID, "create_41")
    create_41_select.send_keys('子宫小于10周')
    # 肌瘤大小
    create_42_select = driver.find_element(By.ID, "create_42")
    create_42_select.send_keys('5-10cm')
    # 按生长部位
    create_45_select = driver.find_element(By.ID, "create_45")
    create_45_select.send_keys('子宫体')
    # 子宫肌瘤的分型（国际妇产科联盟（FIGO）
    create_48_select = driver.find_element(By.ID, "create_48")
    create_48_select.send_keys('Ⅰ型：无蒂黏膜下肌瘤，向肌层扩展≤50%')
    # 是否伴发有全身系统性疾病
    driver.find_element(By.XPATH, '//span[contains(text(), "是否伴发有全身系统性疾病")]/following-sibling::span[@key="n"]').click()

    # UM-2 子宫肌瘤治疗情况
    # 子宫肌瘤治疗方式选择
    create_UM_264_select = driver.find_element(By.ID, "create_UM_264")
    create_UM_264_select.send_keys('子宫肌瘤手术治疗')
    # 手术治疗符合适应证
    create_54_select = driver.find_element(By.ID, "create_54")
    create_54_select.send_keys('手术适应证')
    # 手术适应证选择
    create_57_select = driver.find_element(By.ID, "create_57")
    create_57_select.send_keys('子宫肌瘤合并月经过多或异常出血甚至导致贫血')
    # 是否有手术禁忌证
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术禁忌证")]/following-sibling::span[@key="n"]').click()
    # 是否有术前预处理
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有术前预处理")]/following-sibling::span[@key="n"]').click()
    # 手术路径选择
    create_67_select = driver.find_element(By.ID, "create_67")
    create_67_select.send_keys('经腹手术（开腹术式）')

    # UM-3 预防性抗菌药物应用时机
    # 切口类别
    create_UM_256_select = driver.find_element(By.ID, "create_UM_256")
    create_UM_256_select.send_keys('Ⅱ类切口')
    # 是否使用预防性抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "是否使用预防性抗菌药物")]/following-sibling::span[@key="y"]').click()
    # 预防性抗菌药物选择
    driver.find_element(By.XPATH, '//span[contains(text(), "预防性抗菌药物选择")]/following-sibling::span[@key="d"]').click()
    # 使用首剂抗菌药物起始时间
    driver.execute_script(js, "create_CM_41")
    # delta_minutes = random.randint(20, 40)
    delta = timedelta(minutes = 30)
    create_CM_41_date = surgery_begin_time_obj - delta
    create_CM_41_input = driver.find_element(By.ID, "create_CM_41")
    create_CM_41_input.send_keys(create_CM_41_date.strftime('%Y-%m-%d %H:%M'))
    # 手术时间是否≥3小时
    driver.find_element(By.XPATH, '//span[contains(text(), "手术时间是否≥3小时")]/following-sibling::span[@key="n"]').click()
    # 术中出血量是否≥1500ml
    driver.find_element(By.XPATH, '//span[contains(text(), "术中出血量是否≥1500ml")]/following-sibling::span[@key="n"]').click()
    # 术后抗菌药物停止使用时间
    driver.execute_script(js, "create_CM_47")
    delta_hours = random.randint(24, 48)
    delta = timedelta(hours = delta_hours)
    create_CM_47_date = surgery_end_time_obj + delta
    create_CM_47_input = driver.find_element(By.ID, "create_CM_47")
    create_CM_47_input.send_keys(create_CM_47_date.strftime('%Y-%m-%d %H:%M'))

    # UM-4 输血量
    # 术中腹腔内出血量(ml)
    delta = random.randint(300, 500) // 100 * 100
    create_119_select = driver.find_element(By.ID, "create_119")
    create_119_select.send_keys(delta)
    # 是否实施输血
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施输血")]/following-sibling::span[@key="n"]').click()

    # UM-5 术后并发症与再手术情况
    # 是否有治疗中、治疗后并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有治疗中、治疗后并发症")]/following-sibling::span[@key="n"]').click()

    # UM-6 住院期间为患者提供术前、术后健康教育与出院时提供教育告知五要素情况
    check_pre_post_op_health_education(driver, row)

    # UM-7 手术切口愈合情况
    check_surgical_wound_healing(driver, row)

    # UM-8 离院方式
    check_discharge_method(driver, row)

    # UM-9 患者对服务的体验与评价
    patient_service_evaluation(driver, row)

    # UM-10 住院费用
    basic_fee(driver, row)
    pass

# 宫颈癌（手术治疗）
def disease_cervical_cancer(driver, row):
    #  基本信息
    # 发病日期时间是否无法确定或无记录
    driver.find_element(By.XPATH, '//span[contains(text(), "发病日期时间是否无法确定或无记录")]/following-sibling::span[@key="UTD"]').click()
    basic_info(driver, row)
    # 到达本院急诊或者门诊日期时间是否无法确定或无记录
    driver.find_element(By.XPATH, '//input[@id="create_CM_19"]/following-sibling::span[@key="UTD"]').click()

    # 手术时间
    surgery_begin_time_obj = row['surgery_begin_time_obj']
    surgery_end_time_obj = row['surgery_end_time_obj']
    # 入住ICU日期时间
    driver.execute_script(js, "create_CM_22")
    # delta_hours = random.randint(2, 4)
    # delta = timedelta(hours = delta_hours)
    create_CM_22_date = surgery_end_time_obj
    create_CM_22_input = driver.find_element(By.ID, "create_CM_22")
    create_CM_22_input.send_keys(create_CM_22_date.strftime('%Y-%m-%d %H:%M'))
    # 离开ICU日期时间
    driver.execute_script(js, "create_CM_23")
    delta_hours = random.randint(24, 48)
    delta = timedelta(hours = delta_hours)
    create_CM_23_date = surgery_end_time_obj + delta
    create_CM_23_input = driver.find_element(By.ID, "create_CM_23")
    create_CM_23_input.send_keys(create_CM_23_date.strftime('%Y-%m-%d %H:%M'))

    # CC-1: 术前评估以及FIGO/TNM分期
    # 治疗前是否有病理组织形态学/细胞学诊断报告
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否有病理组织形态学/细胞学诊断报告")]/following-sibling::span[@key="y"]').click()
    # 采集组织或细胞学标本来源途经
    create_CC_6_input = driver.find_element(By.ID, "create_CC_6")
    create_CC_6_input.send_keys('宫颈锥切标本')
    # 是否有FIGO(pTNM) 临床分期结论
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有FIGO(pTNM) 临床分期结论")]/following-sibling::span[@key="n"]').click()
    # 治疗前是否有影像学诊断报告
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否有影像学诊断报告")]/following-sibling::span[@key="y"]').click()
    # 影像学检查项目
    driver.find_element(By.XPATH, '//span[contains(text(), "影像学检查项目")]/following-sibling::span[@key="a"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "影像学检查项目")]/following-sibling::span[@key="b"]').click()
    driver.find_element(By.XPATH, '//span[contains(text(), "影像学检查项目")]/following-sibling::span[@key="c"]').click()
    # 是否有FIGO(rTNM) 临床分期结论
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有FIGO(rTNM) 临床分期结论")]/following-sibling::span[@key="n"]').click()
    # 治疗前是否完成妇科检查
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否完成妇科检查")]/following-sibling::span[@key="y"]').click()
    # 治疗前妇科检查内容
    clickCheckboxes(driver, 'create_CC_24')
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前妇科检查内容")]/following-sibling::span[@key="f"]').click()
    # 治疗前是否完成体能评估
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否完成体能评估")]/following-sibling::span[@key="n"]').click()
    # 治疗前是否完成生化检查
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否完成生化检查")]/following-sibling::span[@key="y"]').click()
    # 生化检查项目
    clickCheckboxes(driver, 'create_CC_32')
    # 治疗前是否完成深静脉栓塞（VTE）风险评估
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否完成深静脉栓塞（VTE）风险评估")]/following-sibling::span[@key="n"]').click()
    # 治疗前是否完成临床FIGO(cTNM) 分期
    driver.find_element(By.XPATH, '//span[contains(text(), "治疗前是否完成临床FIGO(cTNM) 分期")]/following-sibling::span[@key="n"]').click()
    # 是否将替代治疗方案的益处和风险明确告知患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否将替代治疗方案的益处和风险明确告知患者")]/following-sibling::span[@key="y"]').click()
    # 是否将不同手术途径和术式的风险和益处明确告知患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否将不同手术途径和术式的风险和益处明确告知患者")]/following-sibling::span[@key="y"]').click()
    # 是否与患者充分讨论保留生育力的治疗方案
    driver.find_element(By.XPATH, '//span[contains(text(), "是否与患者充分讨论保留生育力的治疗方案")]/following-sibling::span[@key="y"]').click()
    # 是否治疗前接受过MDT会诊的患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否治疗前接受过MDT会诊的患者")]/following-sibling::span[@key="y"]').click()
    # 是否治疗前接受过新辅助化疗的患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否治疗前接受过新辅助化疗的患者")]/following-sibling::span[@key="n"]').click()
    # 是否治疗前接受过介入的患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否治疗前接受过介入的患者")]/following-sibling::span[@key="n"]').click()
    # 是否治疗前接受过根治性放疗的患者
    driver.find_element(By.XPATH, '//span[contains(text(), "是否治疗前接受过根治性放疗的患者")]/following-sibling::span[@key="n"]').click()

    # CC-2: 手术适应证和手术方案
    # 手术适应证选择
    create_CC_61_input = driver.find_element(By.ID, "create_CC_61")
    create_CC_61_input.send_keys('病理提示宫颈癌')
    # 保留生育的手术适应证选择
    driver.find_element(By.XPATH, '//span[contains(text(), "保留生育的手术适应证选择")]/following-sibling::span[@key="a"]').click()
    # 是否有手术禁忌证
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术禁忌证")]/following-sibling::span[@key="n"]').click()
    # 宫颈癌手术方式的选择
    create_CC_68_input = driver.find_element(By.ID, "create_CC_68")
    create_CC_68_input.send_keys('根治性手术')
    # 根治性手术治疗符合原则规范
    driver.find_element(By.XPATH, '//span[contains(text(), "根治性手术治疗符合原则规范")]/following-sibling::span[@key="b"]').click()
    # WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH, '//span[contains(text(), "根治性手术治疗符合原则规范")]/following-sibling::span[@key="b"]'))).click()
    # create_CC_71 = driver.find_element(By.XPATH, '//span[contains(text(), "根治性手术治疗符合原则规范")]/following-sibling::span[@key="b"]')
    # driver.execute_script("arguments[0].scrollIntoView();", create_CC_71)
    # time.sleep(random.uniform(0.3, 0.6))  # 等待页面滚动
    # create_CC_71.click()
    # 是否有手术淋巴结清扫
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术淋巴结清扫")]/following-sibling::span[@key="y"]').click()
    # 淋巴结清扫组别
    driver.find_element(By.XPATH, '//span[contains(text(), "淋巴结清扫组别")]/following-sibling::span[@key="a"]').click()
    # WebDriverWait(driver, 3).until(EC.visibility_of_element_located((By.XPATH, '//span[contains(text(), "淋巴结清扫组别")]/following-sibling::span[@key="a"]'))).click()
    # 淋巴结清扫范围达到层别的结论
    create_CC_80_input = driver.find_element(By.ID, "create_CC_80")
    create_CC_80_input.send_keys('盆腔淋巴结-髂总动脉水平')
    # 前哨淋巴结示踪剂选择
    create_CC_82_input = driver.find_element(By.ID, "create_CC_82")
    create_CC_82_input.send_keys('无法确定,或者无记录')
    # 前哨淋巴结示踪剂注射部位选择
    create_CC_85_input = driver.find_element(By.ID, "create_CC_85")
    create_CC_85_input.send_keys('无法确定,或者无记录')
    # 前哨淋巴结显示部位
    driver.find_element(By.XPATH, '//span[contains(text(), "前哨淋巴结显示部位")]/following-sibling::span[@key="UTD"]').click()
    # 术中探查宫颈癌病变涉及的范围
    driver.find_element(By.XPATH, '//span[contains(text(), "术中探查宫颈癌病变涉及的范围")]/following-sibling::span[@key="a"]').click()
    # 实施的宫颈癌术式的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "实施的宫颈癌术式的选择")]/following-sibling::span[@key="a"]').click()
    # 手术标本剖视情况描述
    driver.find_element(By.XPATH, '//span[contains(text(), "手术标本剖视情况描述")]/following-sibling::span[@key="a"]').click()
    # 手术记录描述宫旁切除类型
    create_CC_98_input = driver.find_element(By.ID, "create_CC_98")
    create_CC_98_input.send_keys('无法确定,或者无记录')
    # 术中是否阴道重建
    driver.find_element(By.XPATH, '//span[contains(text(), "术中是否阴道重建")]/following-sibling::span[@key="n"]').click()
    # 术中是否卵巢悬吊
    driver.find_element(By.XPATH, '//span[contains(text(), "术中是否卵巢悬吊")]/following-sibling::span[@key="n"]').click()
    # 术后其他重建方式的选择
    driver.find_element(By.XPATH, '//span[contains(text(), "术后其他重建方式的选择")]/following-sibling::span[@key="UTD"]').click()
    # 手术记录描述宫旁切除类型
    create_CC_117_input = driver.find_element(By.ID, "create_CC_117")
    create_CC_117_input.send_keys('无法确定,或者无记录')
    # 手术路径选择
    create_CC_120_input = driver.find_element(By.ID, "create_CC_120")
    create_CC_120_input.send_keys('开腹术式')
    # 是否主刀手术者具有相应手术资质
    driver.find_element(By.XPATH, '//span[contains(text(), "是否主刀手术者具有相应手术资质")]/following-sibling::span[@key="y"]').click()
    # 是否有手术中并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术中并发症")]/following-sibling::span[@key="n"]').click()
    # 术中是否更改手术治疗方案
    driver.find_element(By.XPATH, '//span[contains(text(), "术中是否更改手术治疗方案")]/following-sibling::span[@key="n"]').click()

    # CC-3: 预防性抗菌药物应用时机
    # 是否使用预防性抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "是否使用预防性抗菌药物")]/following-sibling::span[@key="y"]').click()
    # 预防性抗菌药物选择
    driver.find_element(By.XPATH, '//span[contains(text(), "第一代或第二代头孢菌素")]').click()
    # 使用首剂抗菌药物起始时间
    driver.execute_script(js, "create_CM_41")
    delta = timedelta(minutes = 30)
    create_CM_41_date = surgery_begin_time_obj - delta
    create_CM_41_input = driver.find_element(By.ID, "create_CM_41")
    create_CM_41_input.send_keys(create_CM_41_date.strftime('%Y-%m-%d %H:%M'))
    # 手术时间是否≥3小时
    driver.find_element(By.XPATH, '//span[contains(text(), "手术时间是否≥3小时")]/following-sibling::span[@key="y"]').click()
    # 是否术中追加抗菌药物
    driver.find_element(By.XPATH, '//span[contains(text(), "是否术中追加抗菌药物")]/following-sibling::span[@key="y"]').click()
    # 术中出血量是否≥1500ml
    driver.find_element(By.XPATH, '//span[contains(text(), "是否术中追加抗菌药物")]/following-sibling::span[@key="n"]').click()
    # 术后抗菌药物停止使用时间
    driver.execute_script(js, "create_CM_47")
    delta_hours = random.randint(24, 48)
    delta = timedelta(hours = delta_hours)
    create_CM_47_date = surgery_end_time_obj + delta
    create_CM_47_input = driver.find_element(By.ID, "create_CM_47")
    create_CM_47_input.send_keys(create_CM_47_date.strftime('%Y-%m-%d %H:%M'))

    # CC-4: 术后病理及pTNM分期
    # 是否有术后病理报告
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有术后病理报告")]/following-sibling::span[@key="y"]').click()
    # 合格的病理报告包括以下主要内容
    driver.find_element(By.XPATH, '//span[contains(text(), "合格的病理报告包括以下主要内容")]/following-sibling::span[@key="a"]').click()
    # 是否有根据术后病理进行pTNM分期
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有根据术后病理进行pTNM分期")]/following-sibling::span[@key="n"]').click()

    # CC-5: 术后综合治疗方案

    # CC-6: 术后并发症及再手术情况
    # 是否有手术后并发症
    driver.find_element(By.XPATH, '//span[contains(text(), "是否有手术后并发症")]/following-sibling::span[@key="n"]').click()
    # 是否是非计划二次手术
    driver.find_element(By.XPATH, '//span[contains(text(), "是否是非计划二次手术")]/following-sibling::span[@key="n"]').click()

    # CC-7: 输血量
    # 术中出血量(ml)
    delta = random.randint(500, 1000) // 100 * 100
    create_CC_182_input = driver.find_element(By.ID, "create_CC_182")
    create_CC_182_input.send_keys(delta)
    # 术后出血量(ml)
    delta = random.randint(100, 200) // 100 * 100
    create_CC_183_input = driver.find_element(By.ID, "create_CC_183")
    create_CC_183_input.send_keys(delta)
    # 是否实施输血
    driver.find_element(By.XPATH, '//span[contains(text(), "是否实施输血")]/following-sibling::span[@key="n"]').click()

    # CC-8: 住院期间为患者提供术前、术后健康教育与出院时提供教育告知五要素情况
    check_pre_post_op_health_education(driver, row)

    # CC-9: 手术切口愈合情况
    check_surgical_wound_healing(driver, row)

    # CC-10: 离院方式
    check_discharge_method(driver, row)

    # CC-11: 患者对服务的体验与评价
    patient_service_evaluation(driver, row)


    # CC-12: 住院费用
    basic_fee(driver, row)
    pass


if __name__=="__main__":
    try: 
        execute()
    except Exception as e:
        print(e)
        print(traceback.format_exc())
    
