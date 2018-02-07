#encoding=utf-8  
#ZPF  
import win32serviceutil
import win32service
import win32event
import cx_Oracle
import shutil
import time



class PythonService(win32serviceutil.ServiceFramework):   
    #服务名  
    _svc_name_ = "KeasyDeleteManager"  
    #服务在windows系统中显示的名称  
    _svc_display_name_ = "KeasyDeleteManager"  
    #服务的描述  
    _svc_description_ = "This code is a Python service Test"  
  
    def __init__(self, args):   
        win32serviceutil.ServiceFramework.__init__(self, args)   
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)  
  
    def SvcDoRun(self):  
        # 把自己的代码放到这里，就OK
        while True:
            import logging
            from logging import handlers
            logging.basicConfig(level=logging.INFO)
            logger = logging.getLogger(__name__)
            handler = logging.FileHandler('KeasyDeleteManager.log')
            handler.setLevel(logging.INFO)
            formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            handler.setFormatter(formatter)
            logger.addHandler(handler)
            logger.info('开始记录日志----------------------------------------')

            conn = cx_Oracle.connect('spectra/artceps@kayisoft2016:1521/spectra')
            c = conn.cursor()
            s = c.execute("select distinct(t.access_no),t.study_key,v.pathname ||'\\'|| replace(t.pathname,'/','\\') as pathname from SV_UIMAGELOC t , volume v where v.volume_code=t.volume_code and t.access_no is not null and t.instance_stat='R'")
            x = s.fetchall()
            for i in x:
                logger.info(i[2])
                shutil.rmtree(i[2])
                logger.info ('已经删除访问号为 %s 的病人,study_key是%d 路径为 %s' %(i[0],i[1],i[2]))
                s2 = c.execute("update instance set instance_stat='D' where study_key= %d " %i[1])
                conn.commit()
                logger.info ('标记数据状态为已删除！')
            c.close()
            conn.close()
            time.sleep(60)
        # 等待服务被停止   
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)   
              
    def SvcStop(self):   
        # 先告诉SCM停止这个过程   
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)   
        # 设置事件   
        win32event.SetEvent(self.hWaitStop)   
  
if __name__=='__main__':   
    win32serviceutil.HandleCommandLine(PythonService)    
    #括号里参数可以改成其他名字，但是必须与class类名一致；
    #如果要安装服务为python DeleteManager.py install
    # 如果要调试服务为python DeleteManager.py debug