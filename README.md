
#Adodb
##新建项目
![1](1.jpg)
![2](2.jpg)
![3](3.jpg)

---
##前期准备
###资源视图
要进行可视化对话框编辑需打开如下资源视图
![33](33.jpg)
打开方式如下图
![34](34.jpg)

---
##添加菜单栏和对话框
###添加菜单栏
在资源视图中双击IDR_MAINFRAME
![40](40.jpg)
在帮助后空白处右键，并点击新插入
![35](35.jpg)
双击新插入的空白处，添加标题
或者在右下角的属性中的描述文字后添加标题
![36](36.jpg)
![37](37.jpg)
在刚才添加完的菜单底下进行同样的操作，插入并修改名称
<font color="red">在属性中最好改一下ID，方便之后操作</font>
![38](38.jpg)
![39](39.jpg)

###添加对话框
在资源视图的Dialog上右键并点击插入Dialog
![41](41.jpg)
在右下角修改，叫啥自己定
描述文字里的内容是显示在对话框的左上角
![42](42.jpg)
![43](43.jpg)
添加所需要的控件
![44](44.jpg)
<font color="red">编辑框和按钮最好改一下ID，方便之后操作
注：列表控件的属性中还需改一下视图才会变成图中的那种效果</font>
![45](45.jpg)

---
##实现点击菜单栏按钮呼出窗口
![4](4.jpg)
![5](5.jpg)
跳转至刚才选择的类列表文件
在OnList中添加代码

    AdodbDlg dlg;  
    dlg.DoModal();

在开头添加头文件

    #include "AdodbDlg.h"

<font color="red">注：将两段代码中的AdodbDlg改为刚才自己添加的类的名称</font>

![6](6.jpg)
![7](7.jpg)

---
##实现数据库连接并在窗口中显示
将森哥的这个测试文件放到资源文件中
![8](8.jpg)
回到刚才创建的Dialog
![9](9.jpg)
![10](10.jpg)
操作完后发现xx.cpp中多了个函数
![13](13.jpg)
然后添加变量
![11](11.jpg)
![12](12.jpg)
在framework.h中添加以下代码

    #import "C:\Program Files\Common Files\System\ado\msado15.dll"   rename("EOF", "adoEOF") using namespace ADODB; 
       		
![14](14.jpg)
在头文件xx.h中添加以下代码(创建数据库连接、操作对象和变量)

    _ConnectionPtr adodbConnection;
    _RecordsetPtr adodbRecordset;
    _variant_t vUsername, vBirthday, vID, vOld;

![15](15.jpg)
回到xx.cpp中
在OnInitDialog()中添加以下代码

    ::SendMessage(userlist.m_hWnd, LVM_SETEXTENDEDLISTVIEWSTYLE, LVS_EX_FULLROWSELECT, LVS_EX_FULLROWSELECT); //高亮
	//////////为列表控件添加列//////////
	userlist.InsertColumn(0, _T("用户ID"), LVCFMT_LEFT, 60);
	userlist.InsertColumn(1, _T("用户名"), LVCFMT_LEFT, 100);
	userlist.InsertColumn(2, _T("年龄"), LVCFMT_LEFT, 60);
	userlist.InsertColumn(3, _T("生日"), LVCFMT_LEFT, 130);

	HRESULT hr;
	try
	{
		hr = adodbConnection.CreateInstance(_T("ADODB.Connection"));///创建Connection对象
		if (SUCCEEDED(hr))
		{
			hr = adodbConnection->Open("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=test.accdb", "", "", adModeUnknown);///连接数据库
		}
	}
	catch (_com_error e)///捕捉异常
	{
		CString errormessage;
		errormessage.Format(_T("连接数据库失败!\r\n错误信息:%s"), e.ErrorMessage());
		AfxMessageBox(errormessage);///显示错误信息
		return FALSE;
	}
	RefreshData();

![16](16.jpg)
添加函数RefreshData()
<font color="red">方便点的方法：在类视图右键名为xx的类，点击添加函数
笨点的方法：在xx.h添加函数定义，然后在xx.cpp添加函数
</font>

    void AdodbDlg::RefreshData()
    {
        int i = 0;
        userlist.DeleteAllItems();
        try
        {
            adodbRecordset.CreateInstance("ADODB.Recordset");
            adodbRecordset->Open("SELECT * FROM users", _variant_t((IDispatch*)adodbConnection, true), adOpenStatic, adLockOptimistic, adCmdText);
            //m_bSuccess = TRUE;
            while (!adodbRecordset->adoEOF)
            {
                vID = adodbRecordset->GetCollect("ID");
                vUsername = adodbRecordset->GetCollect("username");
                vOld = adodbRecordset->GetCollect("old");
                vBirthday = adodbRecordset->GetCollect("birthday");
                userlist.InsertItem(i, (_bstr_t)vID);
                userlist.SetItemText(i, 1, (_bstr_t)vUsername);
                userlist.SetItemText(i, 2, (_bstr_t)vOld);
                userlist.SetItemText(i, 3, (_bstr_t)vBirthday);

                adodbRecordset->MoveNext();
                i++;
            }
        }
        catch (_com_error e)///捕捉异常
        {
            MessageBox(_T("读取数据库失败!"));///显示错误信息
        }
    }

![17](17.jpg)
调试成功！
<font color="red">调试有问题，详见底下问题1和2</font>
![18](18.jpg)

---
##实现添加操作
转到dialog文件
添加变量，4个编辑框都要(以下只演示一个)
![19](19.jpg)
![20](20.jpg)
<font color="red">在类向导中操作也行，类向导中能查看也能添加</font>
![21](21.jpg)
然后
![22](22.jpg)
![23](23.jpg)
转到xx.cpp
在OnBnClickedAdd()中添加以下代码

    SaveData();
    
![25](25.jpg)
跟之前一样添加函数SaveData()，并添加以下代码

    if (UpdateData())
    if (eUsername.GetLength() > 0)
    {
        adodbRecordset->AddNew();
        if (!adodbRecordset->adoEOF)
        {
            vID = (long)eID;
            vUsername = eUsername;
            vOld = (long)eOld;
            vBirthday = eBirthday;
            adodbRecordset->PutCollect("ID", vID);
            adodbRecordset->PutCollect("username", vUsername);
            adodbRecordset->PutCollect("old", vOld);
            adodbRecordset->PutCollect("birthday", vBirthday);
            adodbRecordset->Update();
            RefreshData();
        }
    }

![26](26.jpg)
调试发现成功
![24](24.jpg)

---
##实现选取列表中某行并显示在底下
转到dialog文件
与之前添加操作类似，右键列表控件，点击事件处理程序，将消息类型改为NM_CLICK
![28](28.jpg)
转到xx.cpp
在OnNMClickList1()中添加以下代码

    int m_nCurrentSel = userlist.GetSelectionMark();
	adodbRecordset->Move(m_nCurrentSel, _variant_t((long)adBookmarkFirst));
	vID = adodbRecordset->GetCollect("ID");
	vUsername = adodbRecordset->GetCollect("username");
	vOld = adodbRecordset->GetCollect("old");
	vBirthday = adodbRecordset->GetCollect("birthday");
	eID = vID.lVal;
	eUsername = (LPCTSTR)(_bstr_t)vUsername;
	eOld = vOld.lVal;
	eBirthday = vBirthday;
	UpdateData(FALSE);
![29](29.jpg)
调试发现成功
![30](30.jpg)

---
##实现删除操作
转到dialog文件
与之前添加操作类似，右键删除按钮，点击事件处理程序，然后确定
转到xx.cpp
在OnBnClickedDel()中添加以下代码

    CString strSQL;
	if (MessageBox(_T("确定删除该记录吗？"), _T("删除记录"), MB_OKCANCEL) == IDOK)
	{
		adodbRecordset.CreateInstance("ADODB.Recordset");
		strSQL.Format(_T("delete from users where username='%s'"), eUsername);
		adodbRecordset->Open((_bstr_t)strSQL, _variant_t((IDispatch*)adodbConnection, true), adOpenStatic, adLockOptimistic, adCmdText);

		RefreshData();
	}

![27](27.jpg)
调试发现成功
![31](31.jpg)

---
##实现修改操作
转到dialog文件
与之前添加操作类似，右键修改按钮，点击事件处理程序，然后确定
转到xx.cpp
在OnBnClickedUpdate()中添加以下代码

    int m_nCurrentSel = -1;
	m_nCurrentSel = userlist.GetSelectionMark();
	if (m_nCurrentSel >= 0)
	{
		try
		{

			adodbRecordset->MoveFirst();
			adodbRecordset->Move(long(m_nCurrentSel));
			if (MessageBox(_T("确定修改该记录吗？"), _T("修改记录"), MB_OKCANCEL) == IDOK)
			{
				adodbRecordset->Delete(adAffectCurrent);
				SaveData();
			}

		}
		catch (_com_error e)///捕捉异常
		{
			MessageBox(_T("读取数据库失败!"));///显示错误信息
		}

	}
	else
		MessageBox(_T("请在列表控件中选择要修改的记录"));

调试发现成功
![32](32.jpg)

---
##问题1：连接数据库失败
![bug1](bug1.jpg)
###解决方案
![bug1_1](bug1_1.jpg)
![bug1_2](bug1_2.jpg)

---
##问题2：未定义标识符_ConnectionPtr adodbConnection;等
###解决方案
尝试将framework.h中添加的代码换个位置
![bug2](bug2.jpg)

