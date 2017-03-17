# coding: utf-8

import logging
import os
import platform
import sys
import subprocess
import time
import traceback
from ConfigParser import ConfigParser
from xlwt import *

reload(sys)
sys.setdefaultencoding('utf-8')

logging.basicConfig(
    level=logging.INFO,
    format="[%(asctime)s] %(name)s:%(levelname)s: %(message)s"
)

def get_config(config_file):
    if os.path.exists(config_file) is True:
        config_parser = ConfigParser()
        config_parser.read(config_file)

        package = config_parser.get('config', 'package')
        activity = config_parser.get('config', 'activity')
        times = config_parser.get('config', 'times')
        start_type = config_parser.get('config', 'start_type')
    else:
        print('No {0} file found'.format(config_file))

    return package, activity, times, start_type

def start_adb():
    cmd = "adb devices"
    logging.debug('cmd:{}'.format(cmd))
    if platform.system() == 'Windows':
        subprocess.check_output(cmd, shell=True)
    time.sleep(2)

def OnlyNum(s,oth=''):
    fomart = '0123456789'
    s2 = ''
    for c in s:
        if c in fomart:
            s2 += c
    return s2

def get_device_id():
    '''返回 device id'''
    device_dict = {}
    get_device_id_cmd = "adb devices | findstr /e device"   # /e 对一行的结尾进行匹配

    try:
        output = subprocess.check_output(get_device_id_cmd, shell=True)
        logging.debug('connected devices:\r{0}'.format(output))
    except Exception:
        #traceback.print_exc()
        logging.info('All device lost, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        output = None

    if output is not None:
        output = output.split("\n")
        logging.debug('split connect devices id: {0}'.format(output))
        device_id_list = []
        for device_id in output:
            if device_id == "":
                continue
            device_id = device_id.replace("\tdevice\r", "")
            logging.debug('device_id: {0}'.format(device_id))
            device_id_list.append(device_id)
        logging.debug('device_id_list: {0}'.format(device_id_list))

    return device_id_list

def get_device_model(device_id):
    device_model = ""
    get_device_model_cmd = "adb -s {0} shell getprop ro.product.model".format(device_id)
    logging.debug(get_device_model_cmd)

    try:
        output_model = subprocess.check_output(get_device_model_cmd, shell=True)
        logging.debug("output_model: {0}".format(output_model))
    except Exception:
        logging.error('get device model error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()
        output_model = None

    if output_model is not None:
        output_model = output_model.strip("\r\r\n") # 去除行尾的换行光标
        logging.debug(output_model)
        output_model = output_model.split(' ')
        logging.debug(output_model)

    for i in output_model:
        device_model += i

    return device_model


def backstage_app(device_id, package, activity):
    time.sleep(3)

    startup_app_cmd = "adb -s {0} shell am start -W {1}/{2}".format(device_id, package, activity)

    try:
        output = subprocess.check_output(startup_app_cmd, shell=True)
        logging.debug("output: \n{0}".format(output))

        error_msg = 'Error'
        #如果启动app有错误
        if output.find(error_msg) >= 0:
            logging.error('backstage_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        else:
            back_home(device_id)

    except Exception:
        logging.error('backstage_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()
        output = None


def startup_app(device_id, i, package, activity):
    time.sleep(3)
    logging.info('Start-up {0}'.format(package))

    startup_app_cmd = "adb -s {0} shell am start -W {1}/{2}".format(device_id, package, activity)

    try:
        output = subprocess.check_output(startup_app_cmd, shell=True)
        logging.debug("output: \n{0}".format(output))

        error_msg = 'Error'
        #如果启动app有错误
        if output.find(error_msg) >= 0:
            logging.error('start app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        else:
            wait_time_key='WaitTime:'
            total_time_key='TotalTime:'
            this_time_key='ThisTime:'

            count = "第{}次".format(i)
            sheet1.write(i,0,count)

            #有WaitTime
            if output.find(wait_time_key) >=0:
                flag = output.find(wait_time_key)
                wait_time = OnlyNum(output[flag+10:flag+16])
                logging.info("WaitTime: {0}ms".format(wait_time))
                sheet1.write(i,1,int(wait_time))

            #有TotalTime
            if output.find(total_time_key) >=0:
                flag = output.find(total_time_key)
                total_time = OnlyNum(output[flag+10:flag+16])
                logging.info("TotalTime: {0}ms".format(total_time))
                sheet1.write(i,2,int(total_time))

            #有ThisTime
            if output.find(this_time_key) >=0:
                flag = output.find(this_time_key)
                this_time = OnlyNum(output[flag+9:flag+15])
                logging.info("ThisTime: {0}ms".format(this_time))
                sheet1.write(i,3,int(this_time))


    except Exception:
        logging.error('start app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()
        output = None


def stop_app(device_id, package):
    time.sleep(3)

    stop_app_cmd = 'adb -s {0} shell am force-stop {1}'.format(device_id, package)

    try:
        subprocess.check_output(stop_app_cmd, shell=True)

    except Exception:
        logging.error('stop app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def back_home(device_id):
    time.sleep(3)
    #返回主菜单
    back_home_cmd1 = 'adb -s {0} shell input keyevent 3'.format(device_id)
    #打开浏览器
    back_home_cmd2 = 'adb -s {0} shell input keyevent 64'.format(device_id)

    try:
        subprocess.check_output(back_home_cmd1, shell=True)
        time.sleep(1)
        subprocess.check_output(back_home_cmd2, shell=True)
    except Exception:
        logging.error('back_home error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def uninstall_app(device_id, package):
    time.sleep(3)
    logging.info('Uninstall {0}'.format(package))

    uninstall_app_cmd = 'adb -s {0} shell pm uninstall {1}'.format(device_id, package)

    try:
        output = subprocess.check_output(uninstall_app_cmd, shell=True)

        #校验是否卸载成功，成功console会返回Success
        if output.find('Success') >=0 :
            logging.debug('output: {0}'.format(output))
        else:
            logging.error('uninstall_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
            logging.error('output: {0}'.format(output))

    except Exception:
        logging.error('uninstall_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def install_app(device_id, package, apk_path):
    time.sleep(3)
    logging.info('Install {0}'.format(package))
    
    try:
        install_app_cmd = 'adb -s {0} shell pm install {1}'.format(device_id, apk_path)
        logging.debug('install_app_cmd: {0}'.format(install_app_cmd))
        logging.info(unicode('如果手机端提示【静默安装拦截】，请手动点击【允许】','utf-8').encode('gbk'))

        output = subprocess.check_output(install_app_cmd, shell=True)

        #校验是否安装成功，成功console会返回Success
        if output.find('Success') >=0 :
            logging.debug('output: {0}'.format(output))

            clear_app_cmd = 'adb -s {0} shell pm clear {1}'.format(device_id, package)
            logging.debug('clear_app_cmd: {0}'.format(clear_app_cmd))
            subprocess.check_output(clear_app_cmd, shell=True)

        else:
            logging.error('install_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
            logging.error('output: {0}'.format(output))

    except Exception:
        logging.error('install_app error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def pull_apk_from_device_to_pc(device_id, package):
    get_apk_path_cmd = 'adb -s {0} shell pm path {1}'.format(device_id, package)

    try:
        output = subprocess.check_output(get_apk_path_cmd, shell=True)
        logging.debug('output: {0}'.format(output))
        apk_path = output[8:]
        logging.debug('apk_path: {0}'.format(apk_path))

        pull_apk_cmd = 'adb -s {0} pull {1}'.format(device_id, apk_path)
        output = subprocess.check_output(pull_apk_cmd, shell=True)
        logging.debug('output: {0}'.format(output))

        return apk_path

    except Exception:
        logging.error('pull_apk_from_device_to_pc error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def cp_apk_to_tmpDir(device_id, package):
    get_apk_path_cmd = 'adb -s {0} shell pm path {1}'.format(device_id, package)

    try:
        output = subprocess.check_output(get_apk_path_cmd, shell=True)
        logging.debug('output: {0}'.format(output))
        apk_path = output[8:]
        logging.debug('apk_path: {0}'.format(apk_path))

        tmp_dir = '/data/local/tmp'

        cp_apk_cmd = 'adb -s {0} shell cp {1} {2}'.format(device_id, apk_path, tmp_dir)
        cp_apk_cmd = cp_apk_cmd.replace('\n','')
        logging.debug('cp_apk_cmd: {0}'.format(cp_apk_cmd))

        output = subprocess.check_output(cp_apk_cmd, shell=True)
        logging.debug('output: {0}'.format(output))

        return apk_path, tmp_dir

    except Exception:
        logging.error('cp_apk_to_tmpDir error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def rm_apk_from_tmpDir(device_id, package, tmp_path_apk):
    
    try:
        rm_tmp_apk_cmd = 'adb -s {0} shell rm {1}'.format(device_id, tmp_path_apk)
        subprocess.check_output(rm_tmp_apk_cmd, shell=True)

    except Exception:
        logging.error('rm_apk_from_tmpDir error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))
        traceback.print_exc()


def run_first(times, package, activity, device_id, device_model):

    apk_path, tmp_dir = cp_apk_to_tmpDir(device_id, package)

    #兼容base.apk
    if apk_path.find('base') >= 0:
        tmp_path_apk = tmp_dir+'/base.apk'
        logging.debug('tmp_path_apk: {0}'.format(tmp_path_apk))
    #其他机型
    else:
        flag = apk_path.index(package)#用index()确保flag有值
        apk = apk_path[flag:]
        tmp_path_apk = tmp_dir+'/'+apk
        logging.debug('tmp_path_apk: {0}'.format(tmp_path_apk))

    #循环启动app
    for i in xrange(times):
        logging.info("Schedule: {0}/{1}".format(i+1, times))
        uninstall_app(device_id, package)
        install_app(device_id, package, tmp_path_apk)
        startup_app(device_id, i+1, package, activity)
        stop_app(device_id, package)

    #在execl中统计
    if i == times-1:
        average_wait = "AVERAGE(B2:B{0})".format(times+1)
        average_total = "AVERAGE(C2:C{0})".format(times+1)
        average_this = "AVERAGE(D2:D{0})".format(times+1)
        sheet1.write(times+1,0,"平均值")
        sheet1.write(times+1,1,Formula(average_wait))
        sheet1.write(times+1,2,Formula(average_total))
        sheet1.write(times+1,3,Formula(average_this))
        sheet1.write(times+3,0,"启动方式")
        sheet1.write(times+3,1,"首次启动")


    if tmp_path_apk != '' and tmp_path_apk != '/':
        rm_flag = rm_apk_from_tmpDir(device_id, package, tmp_path_apk)
    else:
        logging.error('error,tmp_path_apk:{0}'.format(tmp_path_apk))



def run_cold(times, package, activity, device_id):

    #确保被测app处于关闭状态
    stop_app(device_id, package)

    #循环启动app
    for i in xrange(times):
        logging.info("Schedule: {0}/{1}".format(i+1, times))
        startup_app(device_id, i+1, package, activity)
        stop_app(device_id, package)

    #在execl中统计
    if i == times-1:
        average_wait = "AVERAGE(B2:B{0})".format(times+1)
        average_total = "AVERAGE(C2:C{0})".format(times+1)
        average_this = "AVERAGE(D2:D{0})".format(times+1)
        sheet1.write(times+1,0,"平均值")
        sheet1.write(times+1,1,Formula(average_wait))
        sheet1.write(times+1,2,Formula(average_total))
        sheet1.write(times+1,3,Formula(average_this))
        sheet1.write(times+3,0,"启动方式")
        sheet1.write(times+3,1,"冷启动")
  

def run_warm(times, package, activity, device_id):

    #确保被测app处于后台运行状态
    backstage_app(device_id, package, activity)

    #循环启动app
    for i in xrange(times):
        logging.info("Schedule: {0}/{1}".format(i+1, times))
        startup_app(device_id, i+1, package, activity)
        back_home(device_id)

    #在execl中统计
    if i == times-1:
        average_wait = "AVERAGE(B2:B{0})".format(times+1)
        average_total = "AVERAGE(C2:C{0})".format(times+1)
        average_this = "AVERAGE(D2:D{0})".format(times+1)
        sheet1.write(times+1,0,"平均值")
        sheet1.write(times+1,1,Formula(average_wait))
        sheet1.write(times+1,2,Formula(average_total))
        sheet1.write(times+1,3,Formula(average_this))
        sheet1.write(times+3,0,"启动方式")
        sheet1.write(times+3,1,"热启动")


if __name__ == '__main__':
    try:
        #获取配置文件
        wkdir = os.getcwd()
        config_file = "{0}\default.conf".format(wkdir)
        package, activity, times, start_type = get_config(config_file)
        logging.info('get config:\nPackage = {0}\nActivity = {1}\nTimes = {2}\nStart_Type = {3}'.format(package, activity, times, start_type))

        #判断配置信息是否正确
        if not times.isdigit() or int(times)>100 or int(times)<=0 or not start_type.isdigit() or int(start_type)>3 or int(start_type)<1:
            raw_input(unicode('配置信息错误，请重新配置default.conf','utf-8').encode('gbk'))
        else:
            #打开execl        
            file = Workbook(encoding = 'utf-8') 
            sheet1 = file.add_sheet('测试结果')
            sheet1.write(0,0,'启动次数')
            sheet1.write(0,1,'WaitTime(ms)')
            sheet1.write(0,2,'TotalTime(ms)')
            sheet1.write(0,3,'ThisTime(ms)')

            logging.info('Initializing...')
            #获取device信息
            start_adb()
            time.sleep(1)

            device_id_list = []
            device_id_list = get_device_id()
            device_id = device_id_list[0]
            logging.debug('current device id: {0}'.format(device_id))
            device_model = get_device_model(device_id)
            logging.debug('current device model: {0}'.format(device_model))

            times = int(times)
            start_type = int(start_type)

            #首次启动
            if start_type == 1:
                run_first(times, package, activity, device_id, device_model)
                test_type = '首次启动'
            #冷启动
            elif start_type == 2:
                run_cold(times, package, activity, device_id)
                test_type = '冷启动'
            #热启动
            elif start_type == 3:
                run_warm(times, package, activity, device_id)
                test_type = '热启动'
                #sheet1.write_merge(0, 0, 1, 3, '热启动')
            else:
                logging.error('start_type error, @{0} line:{1}'.format(__file__, sys._getframe().f_lineno))        

            timestamp = time.strftime('%Y-%m-%d-%H-%M-%S',time.localtime(time.time()))

            filename = "{0}_{1}_{2}.xls".format(device_model, unicode(test_type,'utf-8').encode('gbk'), timestamp)
            file.save(filename)

    except Exception:
        traceback.print_exc()
        time.sleep(1)
        raw_input(unicode('发生错误, 按回车键退出...','utf-8').encode('gbk'))
    finally:
        pass
