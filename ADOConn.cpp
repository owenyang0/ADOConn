
1
2
3
4
5
6
7
8
9
10
11
12
13
14
15
16
17
18
19
20
21
22
23
24
25
26
27
28
29
30
31
32
33
34
35
36
37
38
39
40
41
42
43
44
45
46
47
48
49
50
51
52
53
54
55
56
57
58
59
60
61
62
63
64
65
66
67
68
69
70
71
72
73
74
75
76
77
78
79
80
81
82
83
84
85
86
87
88
89
90
91
92
93
94
95
96
97
98
99
100
101
102
103
104
105
106
107
108
109
110
111
112
113
114
115
116
117
118
119
120
121
122
123
124
125
126
127
128
129
130
131
132
133
134
135
136
137
138
139
140
141
142
143
144
145
146
147
148
149
150
151
ADOConn.h
#if !defined(AFX_ADOCONN_H__56A56674_91DC_43BB_BD09_9A0C8995161E__INCLUDED_)
#define AFX_ADOCONN_H__56A56674_91DC_43BB_BD09_9A0C8995161E__INCLUDED_
 
#import "C:\Program Files\Common Files\System\ado\msado15.dll"no_namespace \
rename("EOF","adoEOF")rename("BOF","adoBOF")
#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000
class ADOConn
{
public:
ADOConn(void);
~ADOConn(void);
 
 
//添加一个指向Connection对象的指针
_ConnectionPtr m_pConnection;
//添加一个指向Recordset对象的指针
_RecordsetPtr m_pRecordset;
//_RecordsetPtr& GetRecordSet(_bstr_t bstrSQL);
/* //封闭ADO类，方便以后使用 */
void OnInitADOConn(void);
void ExitConnect(void);
// 打开记录集
_RecordsetPtr& GetRecordSet(_bstr_t bstrSQL);
};
#endif // !defined(AFX_ADOCONN_H__56A56674_91DC_43BB_BD09_9A0C8995161E__INCLUDED_)
 
ADOConn.cpp
#include "StdAfx.h"
#include "ADOConn.h"
 
 
ADOConn::ADOConn(void)
{
}
 
 
ADOConn::~ADOConn(void)
{
}
 
 
// //封闭ADO类，方便以后使用
void ADOConn::OnInitADOConn(void)
{
::CoInitialize(NULL);
try
{
//创建connection对象
m_pConnection.CreateInstance("ADODB.Connection");
//设置连接字符串
_bstr_t strConnect="uid=;pwd=;DRIVER={Microsoft Access Driver (*.mdb)};DBQ=Grades.mdb;";
//SERVER和UID,PWD的设置根据实际情况来设置"
/*m_pConnection->Open(strConnect,_T("admin"),_T("owenyang"),adModeUnknown);*/
m_pConnection->Open(strConnect,"admin","owenyang",adModeUnknown);
}
catch (_com_error e)
{
//显示错误信息
AfxMessageBox(e.Description());
}
}
 
 
void ADOConn::ExitConnect(void)
{
//关闭记录集和连接
if (m_pRecordset!=NULL)
{
m_pRecordset->Close();
 
}
m_pConnection->Close();
::CoUninitialize();
}
 
 
// 打开记录集
_RecordsetPtr& ADOConn::GetRecordSet(_bstr_t bstrSQL)
{
//TODO: insert return statement here
try
{
if (m_pConnection==NULL)
{
OnInitADOConn();
}
//创建记录对象
m_pRecordset.CreateInstance(__uuidof(Recordset));
//取得表中记录
m_pRecordset->Open(bstrSQL,m_pConnection.GetInterfacePtr(),adOpenDynamic,
adLockOptimistic,adCmdText);
}
catch (_com_error e)
{
e.Description();
}
return m_pRecordset;
}
 
readAccessToList.cpp
void CTypeHistoryDlg::readAccessToList(void)
{
CRect rect;
GetClientRect(&rect);
int gridWidth=(rect.Width()-15)/11;
/*CString tem=_T("试试");
tem.Format(_T("%d"),gridWidth);*/
//MessageBox(tem);
m_historyList.InsertColumn(0,_T("编号"),LVCFMT_CENTER,gridWidth-10);
m_historyList.InsertColumn(1,_T("日期"),LVCFMT_CENTER,gridWidth+25);
m_historyList.InsertColumn(2,_T("段数"),LVCFMT_CENTER,gridWidth+10);
m_historyList.InsertColumn(3,_T("速度"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(4,_T("回改"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(5,_T("击键"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(6,_T("码长"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(7,_T("错字"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(8,_T("字数"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(9,_T("键数"),LVCFMT_CENTER,gridWidth);
m_historyList.InsertColumn(10,_T("用时"),LVCFMT_CENTER,gridWidth);
m_historyList.SetExtendedStyle(LVS_EX_FLATSB
|LVS_EX_FULLROWSELECT
|LVS_SHOWSELALWAYS
|LVS_EX_GRIDLINES);
ADOConn m_AdoConn;
m_AdoConn.OnInitADOConn();
CString sql;
sql.Format(_T("select* from grade"));
_RecordsetPtr m_pRecordset;
m_pRecordset = m_AdoConn.GetRecordSet((_bstr_t)sql);
while(m_AdoConn.m_pRecordset->adoEOF==0)
{
m_historyList.InsertItem(0,_T(""));
m_historyList.SetItemText(0,1,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("date"));
m_historyList.SetItemText(0,2,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("para"));
m_historyList.SetItemText(0,3,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("speed"));
m_historyList.SetItemText(0,4,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("back"));
m_historyList.SetItemText(0,5,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("hitkey"));
m_historyList.SetItemText(0,6,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("keylong"));
m_historyList.SetItemText(0,7,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("wronwor"));
m_historyList.SetItemText(0,8,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("wordscount"));
m_historyList.SetItemText(0,9,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("keycount"));
m_historyList.SetItemText(0,10,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("time"));
m_historyList.SetItemText(0,0,(TCHAR*)(_bstr_t)m_pRecordset->GetCollect("id"));
m_pRecordset->MoveNext();
}
m_historyList.SetItemState(0,LVIS_SELECTED|LVIS_FOCUSED,LVIS_SELECTED|LVIS_FOCUSED);
m_AdoConn.ExitConnect();
}