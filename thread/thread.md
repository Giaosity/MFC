<!-- vscode-markdown-toc -->
* 1. [ 用全局变量实现线程创建与终止](#-用全局变量实现线程创建与终止)
* 2. [ 实现事件线程创建与终止](#-实现事件线程创建与终止)

<!-- vscode-markdown-toc-config
	numbering=true
	autoSave=true
	/vscode-markdown-toc-config -->
<!-- /vscode-markdown-toc -->

# Thread
---
##  1. <a name='-用全局变量实现线程创建与终止'></a> 用全局变量实现线程创建与终止
回到资源视图中Menu，打开**IDR_MAINFRAME**

跟之前一样先创建菜单栏，分别更改ID

![1](image/1.jpg)

然后分别右键，添加事件处理程序，在类列表中选择类(以下只演示一个)

![2](image/2.jpg)

![3](image/3.jpg)

回到刚才**所选择的类**中

在开头添加全局变量threadController

    volatile int threadController;

![4](image/4.jpg)

然后在类视图中右键所选择的类，添加以下变量

![5](image/5.jpg)

![6](image/6.jpg)

![7](image/7.jpg)

转到在所选择的类的**头文件**中，可以发现变量添加成功

![8](image/8.jpg)

继续**回到所选择的类**中

在构造函数中添加以下代码

    m_strMessage = "没有线程启动";
	m_iTime = 0;

![9](image/9.jpg)

找到OnDraw()函数中添加以下代码，别忘了把参数中的注释去掉

![10](image/10.jpg)

然后在刚才添加的两个事件处理程序的函数中分别添加以下代码

    //OnStart()下
    threadController = 1;
	HWND hWnd = GetSafeHwnd();
	AfxBeginThread(ThreadProc, hWnd, THREAD_PRIORITY_NORMAL);

    //OnStop()下
    threadController = 0;

![11](image/11.jpg)

在文件开头添加头文件

    #include "MainFrm.h"

![20](image/20.jpg)

最后再添加一个线程函数

    UINT ThreadProc(LPVOID param)
    {
        CMy2020618007高思源App* pApp = (CMy2020618007高思源App*)AfxGetApp();
	    CMainFrame* pMainFrame = (CMainFrame*)pApp->GetMainWnd();
	    CMy2020618007高思源View* pView = (CMy2020618007高思源View*)pMainFrame->GetActiveView();pMainFrame->GetActiveView();
        pView->m_strMessage = "启动了一个线程！";
        while (threadController)
        {	
            ::Sleep(1000);
            pView->m_iTime++;
            pView->Invalidate();
        }
        pView->m_iTime = 0;
        pView->m_strMessage = "线程结束！";
        return 0;
    }

![12](image/12.jpg)

调试发现成功

![13](image/13.jpg)

![14](image/14.jpg)

![15](image/15.jpg)

---

##  2. <a name='-实现事件线程创建与终止'></a> 实现事件线程创建与终止
还是回到资源视图中Menu，打开**IDR_MAINFRAME**

创建菜单栏，分别更改ID，并添加事件处理程序(就不演示了)

![16](image/16.jpg)

**回到所选类的有文件中**，添加消息定义

    const int WM_THREAD_SENDMESS = WM_USER + 20;

![17](image/17.jpg)

添加消息函数声明

    afx_msg LRESULT  OnThreadSendmess(WPARAM w, LPARAM l);

![18](image/18.jpg)

**转到所选的类中**

添加消息映射

    ON_MESSAGE(WM_THREAD_SENDMESS, OnThreadSendmess)

![19](image/19.jpg)

添加消息函数

    LRESULT CMy2020618007高思源View::OnThreadSendmess(WPARAM w, LPARAM l)
    {
        Invalidate();
        return 0;
    }

![21](image/21.jpg)

在开头添加头文件和两个全局变量

    #include "afxmt.h"


    CEvent threadStart;
    CEvent threadEnd;

![22](image/22.jpg)

添加一个线程函数

    UINT ThreadProcEvent(LPVOID param)
    {
        CMy2020618007高思源App* pApp = (CMy2020618007高思源App*)AfxGetApp();
        CMainFrame* pMainFrame = (CMainFrame*)pApp->GetMainWnd();
        CMy2020618007高思源View* pView = (CMy2020618007高思源View*)pMainFrame->GetActiveView();
        ::WaitForSingleObject(threadStart.m_hObject, INFINITE);
        pView->m_strMessage = "启动了一个线程！";
        BOOL keepRunning = TRUE;
        while (keepRunning)
        {
            ::Sleep(1000);
            int result = ::WaitForSingleObject(threadEnd.m_hObject, 0);
            if (result == WAIT_OBJECT_0)
                keepRunning = FALSE;
            pView->m_iTime++;
            ::PostMessage((HWND)param, WM_THREAD_SENDMESS, 0, 0);
        }
        pView->m_iTime = 0;
        pView->m_strMessage = "线程结束！";
        return 0;
    }

分别在添加的事件处理程序函数下添加以下代码

    //OnEventStart()下
    HWND hWnd = GetSafeHwnd();
	AfxBeginThread(ThreadProcEvent, hWnd, THREAD_PRIORITY_NORMAL);
	threadStart.SetEvent();

    //OnEventStop()下
    threadEnd.SetEvent();

![23](image/23.jpg)

调试发现成功

![13](image/13.jpg)

![14](image/14.jpg)

![15](image/15.jpg)

---


