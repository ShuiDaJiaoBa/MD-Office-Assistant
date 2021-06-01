import wx
import re
import os
import json
import pandas as pd
import poplib
from email.parser import Parser
from email.header import decode_header


global userjson

global model_path
global excel_file_folder_path
global excel_output_file_folder_path
global email_file_folder_path
m_checkList4Choices = []
checkedItems = []
col_choose_list = []

global frame1#EXCEL_UPLOAD
global frame2#EXCEL_ADD
global frame3#EMAIL_DOWNLOAD

class Main ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = u"Microsoft Daddy", pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer2 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText1 = wx.StaticText( self, wx.ID_ANY, u"欢迎使用Microsoft Daddy办公助手", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText1.Wrap( -1 )

        bSizer2.Add( self.m_staticText1, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        gSizer3 = wx.GridSizer( 3, 3, 0, 0 )

        self.email_button = wx.Button( self, wx.ID_ANY, u"Email附件下载", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.email_button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.excel_button = wx.Button( self, wx.ID_ANY, u"EXCEL合并", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer3.Add( self.excel_button, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        # self.testdata_button = wx.Button(self, wx.ID_ANY, u"考核数据处理", wx.DefaultPosition, wx.DefaultSize, 0)
        # gSizer3.Add(self.testdata_button, 0, wx.ALL, 5)

        bSizer2.Add( gSizer3, 1, wx.EXPAND, 5 )

        self.about_button = wx.Button( self, wx.ID_ANY, u"关于", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer2.Add( self.about_button, 0, wx.ALL|wx.ALIGN_RIGHT, 5 )


        self.SetSizer( bSizer2 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.email_button.Bind( wx.EVT_BUTTON, self.Open_function_email )
        self.excel_button.Bind( wx.EVT_BUTTON, self.Open_function_excel )
        self.about_button.Bind( wx.EVT_BUTTON, self.open_about_message )

    def __del__( self ):
        pass


    # Virtual event handlers, overide them in your derived class
    def Open_function_email( self, event ):
        self.email_button.Enable(False)
        global frame3
        frame3 = EMAIL_DOWNLOAD(None)
        frame3.Show()

    def Open_function_excel( self, event ):
        self.excel_button.Enable(False)
        global frame4
        frame4 = EXCEL_UPLOAD(None)
        frame4.Show()

    def open_about_message( self, event ):
        self.about = About(None)
        self.about.Show()

class About ( wx.Dialog ):

    def __init__( self, parent ):
        wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer3 = wx.BoxSizer( wx.VERTICAL )

        self.about1 = wx.StaticText( self, wx.ID_ANY, u"作者：王思薇", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.about1.Wrap( -1 )

        bSizer3.Add( self.about1, 0, wx.ALL, 5 )

        self.about2 = wx.StaticText( self, wx.ID_ANY, u"邮箱：caozhengsiwei@163.com", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.about2.Wrap( -1 )

        bSizer3.Add( self.about2, 0, wx.ALL, 5 )

        self.about3 = wx.StaticText( self, wx.ID_ANY, u"Version：1.2", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.about3.Wrap( -1 )

        bSizer3.Add( self.about3, 0, wx.ALL, 5 )


        self.SetSizer( bSizer3 )
        self.Layout()
        bSizer3.Fit( self )

        self.Centre( wx.BOTH )

    def __del__( self ):
        pass

class EMAIL_DOWNLOAD ( wx.Frame ):
    def __init__(self, parent):
        wx.Frame.__init__(self, parent, id=wx.ID_ANY, title=wx.EmptyString, pos=wx.DefaultPosition,
                          size=wx.Size(604, 341), style=wx.DEFAULT_FRAME_STYLE | wx.TAB_TRAVERSAL)

        self.SetSizeHints(wx.DefaultSize, wx.DefaultSize)

        bSizer4 = wx.BoxSizer(wx.VERTICAL)

        self.email_title = wx.StaticText(self, wx.ID_ANY, u"Email附件下载", wx.DefaultPosition, wx.DefaultSize, 0)
        self.email_title.Wrap(-1)

        bSizer4.Add(self.email_title, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_staticText7 = wx.StaticText(self, wx.ID_ANY, u"##说明POP3服务##", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText7.Wrap(-1)

        bSizer4.Add(self.m_staticText7, 0, wx.ALL, 5)

        self.m_checkBox1 = wx.CheckBox(self, wx.ID_ANY, u"自动保存信息", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.m_checkBox1, 0, wx.ALL, 5)

        gSizer6 = wx.GridSizer(2, 2, 0, 0)

        self.m_staticText8 = wx.StaticText(self, wx.ID_ANY, u"邮箱", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText8.Wrap(-1)

        gSizer6.Add(self.m_staticText8, 0, wx.ALL, 5)

        self.m_textCtrl5 = wx.TextCtrl(self, wx.ID_ANY, userjson['email_address'], wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer6.Add(self.m_textCtrl5, 0, wx.ALL, 5)

        self.m_staticText9 = wx.StaticText(self, wx.ID_ANY, u"授权码", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText9.Wrap(-1)

        gSizer6.Add(self.m_staticText9, 0, wx.ALL, 5)

        self.m_textCtrl6 = wx.TextCtrl(self, wx.ID_ANY, userjson['email_password'], wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer6.Add(self.m_textCtrl6, 0, wx.ALL, 5)

        bSizer4.Add(gSizer6, 1, wx.EXPAND, 5)

        gSizer7 = wx.GridSizer(2, 4, 0, 0)

        self.m_staticText11 = wx.StaticText(self, wx.ID_ANY, u"pop服务器", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText11.Wrap(-1)

        gSizer7.Add(self.m_staticText11, 0, wx.ALL, 5)

        self.m_textCtrl8 = wx.TextCtrl(self, wx.ID_ANY, userjson['pop_server_host'], wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer7.Add(self.m_textCtrl8, 0, wx.ALL, 5)

        self.m_staticText13 = wx.StaticText(self, wx.ID_ANY, u"监听端口", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText13.Wrap(-1)

        gSizer7.Add(self.m_staticText13, 0, wx.ALL, 5)

        self.m_spinCtrl3 = wx.SpinCtrl(self, wx.ID_ANY, u"995", wx.DefaultPosition, wx.DefaultSize,
                                       wx.SP_ARROW_KEYS, 0, 10000, 0)
        gSizer7.Add(self.m_spinCtrl3, 0, wx.ALL, 5)

        self.m_staticText16 = wx.StaticText(self, wx.ID_ANY, u"匹配关键字", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText16.Wrap(-1)

        gSizer7.Add(self.m_staticText16, 0, wx.ALL, 5)

        self.m_textCtrl11 = wx.TextCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer7.Add(self.m_textCtrl11, 0, wx.ALL, 5)

        self.m_staticText17 = wx.StaticText(self, wx.ID_ANY, u"检索数目", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText17.Wrap(-1)

        gSizer7.Add(self.m_staticText17, 0, wx.ALL, 5)

        self.m_spinCtrl4 = wx.SpinCtrl(self, wx.ID_ANY, wx.EmptyString, wx.DefaultPosition, wx.DefaultSize,
                                       wx.SP_ARROW_KEYS, 0, 1000, 0)
        gSizer7.Add(self.m_spinCtrl4, 0, wx.ALL, 5)

        bSizer4.Add(gSizer7, 1, wx.EXPAND, 5)

        gSizer9 = wx.GridSizer(0, 3, 0, 0)

        self.m_staticText18 = wx.StaticText(self, wx.ID_ANY, u"请选择输出文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText18.Wrap(-1)

        gSizer9.Add(self.m_staticText18, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_button8 = wx.Button(self, wx.ID_ANY, u"选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer9.Add(self.m_button8, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.m_staticText19 = wx.StaticText(self, wx.ID_ANY, u"未选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText19.Wrap(-1)

        gSizer9.Add(self.m_staticText19, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        bSizer4.Add(gSizer9, 1, wx.EXPAND, 5)

        self.m_button7 = wx.Button(self, wx.ID_ANY, u"开始下载", wx.DefaultPosition, wx.DefaultSize, 0)
        bSizer4.Add(self.m_button7, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)

        self.SetSizer(bSizer4)
        self.Layout()

        self.Centre(wx.BOTH)

        # Connect Events
        self.m_button8.Bind(wx.EVT_BUTTON, self.open_output_filefolder)
        self.m_button7.Bind(wx.EVT_BUTTON, self.Start_download)

        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def __del__(self):
        pass


    # Virtual event handlers, overide them in your derived class
    def open_output_filefolder(self, event):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.m_staticText19.SetLabel(dlg.GetPath().split('\\')[-1])
            global email_file_folder_path
            email_file_folder_path = dlg.GetPath()
            dlg.Destroy()

    def Start_download( self, event ):
        para1 = self.m_textCtrl5.GetValue()
        para2 = self.m_textCtrl6.GetValue()
        para3 = self.m_textCtrl8.GetValue()
        para4 = self.m_spinCtrl3.GetValue()
        para5 = self.m_textCtrl11.GetValue()
        para6 = self.m_spinCtrl4.GetValue()

        self.m_textCtrl5.Enable(False)
        self.m_textCtrl6.Enable(False)
        self.m_textCtrl8.Enable(False)
        self.m_spinCtrl3.Enable(False)
        self.m_textCtrl11.Enable(False)
        self.m_spinCtrl4.Enable(False)

        if (para1 == '') | (para2 == '') | (para3 == '') | (para5 == '') | (self.m_staticText19.GetLabel()=="未选择文件夹"):
            dlg = wx.MessageDialog(None, '请输入完整信息！', u"ERROR", )
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
            self.m_textCtrl5.Enable(True)
            self.m_textCtrl6.Enable(True)
            self.m_textCtrl8.Enable(True)
            self.m_spinCtrl3.Enable(True)
            self.m_textCtrl11.Enable(True)
            self.m_spinCtrl4.Enable(True)
            dlg.Destroy()

        else:
            message = recv_email_by_pop3(para1, para2, para3, para4, para5, para6, email_file_folder_path)
            if message == 'success':
                dlg = wx.MessageDialog(None, '附件下载成功！', u"SUCCESS", )
                if dlg.ShowModal() == wx.ID_YES:
                    self.Close(True)
                self.m_textCtrl5.Enable(True)
                self.m_textCtrl6.Enable(True)
                self.m_textCtrl8.Enable(True)
                self.m_spinCtrl3.Enable(True)
                self.m_textCtrl11.Enable(True)
                self.m_spinCtrl4.Enable(True)

                if self.m_checkBox1.GetValue() == True:
                    userjson['email_address'] = self.m_textCtrl5.GetValue()
                    userjson['email_password'] = self.m_textCtrl6.GetValue()
                    userjson['pop_server_host'] = self.m_textCtrl8.GetValue()
                    with open('userinformation.json', 'w') as f:
                        json.dump(userjson, f)
                    #print('写入成功')

                dlg.Destroy()
            else:
                dlg = wx.MessageDialog(None, message, u"ERROR", )
                if dlg.ShowModal() == wx.ID_YES:
                    self.Close(True)
                self.m_textCtrl5.Enable(True)
                self.m_textCtrl6.Enable(True)
                self.m_textCtrl8.Enable(True)
                self.m_spinCtrl3.Enable(True)
                self.m_textCtrl11.Enable(True)
                self.m_spinCtrl4.Enable(True)
                dlg.Destroy()

    def OnClose(self, event):
        frame0.email_button.Enable(True)
        self.Destroy()

class Dialog ( wx.Dialog ):

    def __init__( self, parent ):
        wx.Dialog.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.DefaultSize, style = wx.DEFAULT_DIALOG_STYLE )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer4 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText7 = wx.StaticText( self, wx.ID_ANY, u"正在处理中，请稍等。。。", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText7.Wrap( -1 )

        bSizer4.Add( self.m_staticText7, 0, wx.ALL, 5 )


        self.SetSizer( bSizer4 )
        self.Layout()
        bSizer4.Fit( self )

        self.Centre( wx.BOTH )

    def __del__( self ):
        pass

class EXCEL_UPLOAD( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = 'EXCEL合并', pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.TAB_TRAVERSAL )

        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )

        bSizer1 = wx.BoxSizer( wx.VERTICAL )

        self.m_staticText_excel_welcome = wx.StaticText( self, wx.ID_ANY, u"EXCEL合并", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText_excel_welcome.Wrap( -1 )

        bSizer1.Add( self.m_staticText_excel_welcome, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )

        self.m_staticText_excel_Maunal = wx.StaticText( self, wx.ID_ANY, u"##说明##", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText_excel_Maunal.Wrap( -1 )

        bSizer1.Add( self.m_staticText_excel_Maunal, 0, wx.ALL, 5 )

        gSizer1 = wx.GridSizer( 3, 2, 0, 0 )

        self.m_button_excel_choose_model = wx.Button( self, wx.ID_ANY, u"选择模板", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_button_excel_choose_model, 0, wx.ALL, 5 )

        self.m_staticText_excel_model_name = wx.StaticText( self, wx.ID_ANY, u"未选择文件", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText_excel_model_name.Wrap( -1 )

        gSizer1.Add( self.m_staticText_excel_model_name, 0, wx.ALL, 5 )

        self.m_button_excel_choose_filefolder = wx.Button( self, wx.ID_ANY, u"选择待合并文件夹", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer1.Add( self.m_button_excel_choose_filefolder, 0, wx.ALL, 5 )

        self.m_staticText_excel_filefolder_name = wx.StaticText( self, wx.ID_ANY, u"未选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText_excel_filefolder_name.Wrap( -1 )

        gSizer1.Add( self.m_staticText_excel_filefolder_name, 0, wx.ALL, 5 )

        self.m_button_excel_choose_output_filefolder = wx.Button(self, wx.ID_ANY, u"选择输出文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer1.Add(self.m_button_excel_choose_output_filefolder, 0, wx.ALL, 5)

        self.m_staticText_excel_output_filefolder_name = wx.StaticText(self, wx.ID_ANY, u"未选择文件夹", wx.DefaultPosition, wx.DefaultSize, 0)
        self.m_staticText_excel_output_filefolder_name.Wrap(-1)

        gSizer1.Add(self.m_staticText_excel_output_filefolder_name, 0, wx.ALL, 5)

        bSizer1.Add( gSizer1, 1, wx.EXPAND, 5 )

        self.upload = wx.Button( self, wx.ID_ANY, u"上传", wx.DefaultPosition, wx.DefaultSize, 0 )
        bSizer1.Add( self.upload, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )


        self.SetSizer( bSizer1 )
        self.Layout()

        self.Centre( wx.BOTH )

        # Connect Events
        self.m_button_excel_choose_model.Bind( wx.EVT_BUTTON, self.choose_model )
        self.m_button_excel_choose_filefolder.Bind( wx.EVT_BUTTON, self.choose_filefolder )
        self.m_button_excel_choose_output_filefolder.Bind( wx.EVT_BUTTON, self.choose_output_filefolder)
        self.upload.Bind( wx.EVT_BUTTON, self.upload_and_progress )

        self.Bind(wx.EVT_CLOSE, self.OnClose)


    def __del__( self ):
        pass


    # Virtual event handlers, overide them in your derived class
    def choose_model( self, event ):
        wildcard = 'All files(*.*)|*.*'
        fileDialog = wx.FileDialog(self, message ="选择单个文件", wildcard = wildcard, style = wx.FD_OPEN)
        if fileDialog.ShowModal() == wx.ID_OK:
            #print(fileDialog.GetPath().split('\\')[-1])
            self.m_staticText_excel_model_name.SetLabel(fileDialog.GetPath().split('\\')[-1])
            global model_path
            model_path = fileDialog.GetPath()
            fileDialog.Destroy()

    def choose_filefolder( self, event ):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.m_staticText_excel_filefolder_name.SetLabel(dlg.GetPath().split('\\')[-1])
            global excel_file_folder_path
            excel_file_folder_path = dlg.GetPath()
            dlg.Destroy()

    def choose_output_filefolder( self, event ):
        dlg = wx.DirDialog(self, u"选择文件夹", style=wx.DD_DEFAULT_STYLE)
        if dlg.ShowModal() == wx.ID_OK:
            self.m_staticText_excel_output_filefolder_name.SetLabel(dlg.GetPath().split('\\')[-1])
            global excel_output_file_folder_path
            excel_output_file_folder_path = dlg.GetPath()
            dlg.Destroy()

    def upload_and_progress(self, event):
        if (self.m_staticText_excel_model_name.GetLabel()!="未选择文件")&(self.m_staticText_excel_filefolder_name.GetLabel()!="未选择文件夹")&(self.m_staticText_excel_output_filefolder_name.GetLabel()!="未选择文件夹"):
            self.m_button_excel_choose_model.Enable(False)
            self.m_button_excel_choose_filefolder.Enable(False)
            self.m_button_excel_choose_output_filefolder.Enable(False)
            self.upload.Enable(False)

            count_sheet_numbet = len(pd.ExcelFile(model_path).sheet_names)
            #print(count_sheet_numbet)

            col_name_list = []
            for i in pd.ExcelFile(model_path).sheet_names:
                col_name_list.append('sheet：'+str(i)+'（必勾！）')
                linshi = pd.DataFrame(pd.read_excel(model_path, sheet_name=i))
                col_name_list.extend(list(linshi))
            m_checkList4Choices.extend(col_name_list)
            global frame2
            frame2 = EXCEL_ADD(None)
            frame2.Show()
        else:
            self.upload.Enable(False)
            dlg = wx.MessageDialog(None, u"请必须选择文件和文件夹！！！", u"ERROR",)
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
            self.upload.Enable(True)
            dlg.Destroy()

    def OnClose(self, event):
        frame0.excel_button.Enable(True)
        self.Destroy()

class EXCEL_ADD ( wx.Frame ):

    def __init__( self, parent ):
        wx.Frame.__init__ ( self, parent, id = wx.ID_ANY, title = wx.EmptyString, pos = wx.DefaultPosition, size = wx.Size( 500,300 ), style = wx.DEFAULT_FRAME_STYLE|wx.STAY_ON_TOP|wx.TAB_TRAVERSAL )
        self.SetSizeHints( wx.DefaultSize, wx.DefaultSize )
        bSizer3 = wx.BoxSizer( wx.VERTICAL )
        self.m_staticText_choose_need_col = wx.StaticText( self, wx.ID_ANY, u"请选择需要合并的列", wx.DefaultPosition, wx.DefaultSize, 0 )
        self.m_staticText_choose_need_col.Wrap( -1 )
        bSizer3.Add( self.m_staticText_choose_need_col, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
        #gSizer2 = wx.GridSizer( 1, 3, 0, 0 )
        gSizer2 = wx.FlexGridSizer(1, 3, 0, 0)
        gSizer2.SetFlexibleDirection(wx.BOTH)
        gSizer2.SetNonFlexibleGrowMode(wx.FLEX_GROWMODE_SPECIFIED)
        self.m_checkList4 = wx.CheckListBox( self, wx.ID_ANY, wx.DefaultPosition, wx.DefaultSize, m_checkList4Choices, 0 )
        gSizer2.Add( self.m_checkList4, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
        self.m_checkBox2 = wx.CheckBox(self, 1, u"全选", wx.DefaultPosition, wx.DefaultSize, 0)
        gSizer2.Add(self.m_checkBox2, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        self.m_button_col_add = wx.Button( self, wx.ID_ANY, u"确定", wx.DefaultPosition, wx.DefaultSize, 0 )
        gSizer2.Add( self.m_button_col_add, 0, wx.ALL|wx.ALIGN_CENTER_HORIZONTAL, 5 )
        bSizer3.Add( gSizer2, 1, wx.EXPAND, 5 )
        self.SetSizer( bSizer3 )
        self.Layout()
        self.Centre( wx.BOTH )
        # Connect Events
        self.Bind(wx.EVT_CHECKBOX, self.choose_or_no, id=1)
        self.m_button_col_add.Bind( wx.EVT_BUTTON, self.col_add_enter )

        self.Bind(wx.EVT_CLOSE, self.OnClose)

    def __del__( self ):
        pass

    def OnClose(self, event):
        frame1.m_button_excel_choose_model.Enable(True)
        frame1.m_button_excel_choose_filefolder.Enable(True)
        frame1.m_button_excel_choose_output_filefolder.Enable(True)
        frame1.upload.Enable(True)
        m_checkList4Choices.clear()
        self.Destroy()

    # Virtual event handlers, overide them in your derived class
    def choose_or_no(self, event):
        #print(self.m_checkBox2.GetValue(), self.m_checkList4.IsChecked(0), self.m_checkList4.IsChecked(1), self.m_checkList4.IsChecked(2))
        all = list(range(0,self.m_checkList4.GetCount()))
        none = list()
        if self.m_checkBox2.GetValue() == True:
            self.m_checkList4.SetCheckedItems(all)
        else:
            self.m_checkList4.SetCheckedItems(none)
        #print('m_checkList4:', self.m_checkList4)

    def col_add_enter( self, event ):
        checkedItems = [i for i in range(self.m_checkList4.GetCount()) if self.m_checkList4.IsChecked(i)]
        # print('m_checkList4Choices:', m_checkList4Choices)
        # print('checkedItems:', checkedItems)
        # print('len(checkedItems):',len(checkedItems))

        count_all = 0
        for i in range(len(m_checkList4Choices)):
            select = re.search('（必勾！）', m_checkList4Choices[i])
            if select != None:
                count_all = count_all + 1
        #print('count_all', count_all)

        count_check = 0
        for i in checkedItems:
            select = re.search('（必勾！）', m_checkList4Choices[i])
            if select != None:
                count_check = count_check + 1
        #print('count_check', count_check)

        if count_all == count_check:
            for i in range(len(checkedItems)):
                col_choose_list.append(m_checkList4Choices[int(checkedItems[i])])
            print(col_choose_list)
            frame1.m_button_excel_choose_model.Enable(True)
            frame1.m_button_excel_choose_filefolder.Enable(True)
            frame1.m_button_excel_choose_output_filefolder.Enable(True)
            frame1.upload.Enable(True)
            self.m_button_col_add.Enable(False)
            self.warnning = Dialog(None)
            self.warnning.Show()
            excel_final_add(model_path, excel_file_folder_path, excel_output_file_folder_path, col_choose_list)##################################################
            self.warnning.Destroy()
            self.Destroy()
            dlg = wx.MessageDialog(None, u"处理成功！！！", u"SUCCESS", )
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
            self.m_button_col_add.Enable(True)
            m_checkList4Choices.clear()
            checkedItems.clear()
            col_choose_list.clear()
            dlg.Destroy()
        else:
            self.m_button_col_add.Enable(False)
            dlg = wx.MessageDialog(None, u"请勾选必须要勾选的列！！！", u"ERROR",)
            if dlg.ShowModal() == wx.ID_YES:
                self.Close(True)
            self.m_button_col_add.Enable(True)
            dlg.Destroy()

def excel_final_add(model, input, output, list):
    # 用ExcelWriter打开excel文件，引擎设置为'openpyxl'以便于保存多个sheet
    outputfilename = output + '\\' + model.split('\\')[-1].split('.')[0] + '_ALL.' + model.split('\\')[-1].split('.')[1]
    trytry = pd.DataFrame()
    trytry.to_excel(outputfilename)
    writer = pd.ExcelWriter(outputfilename, engine='openpyxl')
    # sheet1处理
    # 将excel文件的第一个sheet导入dataframe

    # os.getcwd()为获取当前目录（D:\PycharmProjects\WSW）
    # os.getcwd()+'\接受邮件'为WSW的子目录：接收邮件
    # 下一行为获取当前目录下的所有文件名，包含子目录的
    count_sheet_number = len(pd.ExcelFile(model_path).sheet_names)
    count_sheet_list = pd.ExcelFile(model_path).sheet_names
    # print(list)
    col_to_add = [[] for i in range(count_sheet_number)]
    list_about_col_index = []
    #print('list:', list)
    count_wsw = 0
    for items in list:
        select = re.search('（必勾！）', items)
        if select != None:
            list_about_col_index.append(count_wsw)
        count_wsw = count_wsw + 1
    list_about_col_index.append(count_wsw)
    #print(list_about_col_index)
    count_test = 0
    for i in range(len(list_about_col_index)-1):
        #print('i',i)
        while count_test<list_about_col_index[i+1]:
            if count_test != list_about_col_index[i]:
                col_to_add[i].append(count_test)
            if count_test == list_about_col_index[i+1]:
                continue
            count_test = count_test + 1
    # print(col_to_add)

    for sheet_number in range(count_sheet_number):
        print('[Progressing sheet]:'+count_sheet_list[sheet_number]+'...')
        MODEL_df = pd.DataFrame(pd.read_excel(model, sheet_name=count_sheet_list[sheet_number]))
        for root, dirs, files in os.walk(input):
            # files是一个list，包含所有文件名
            for name in files:
                print('--Filename:', name)
                # 剔除模板excel和最终输出的excel
                if (name != model.split('\\')[-1]) & (name != outputfilename.split('\\')[-1]):
                    # 设定需要处理的文件路径
                    EXCELpath = input + '\\' + name
                    #print('EXCELpath:', EXCELpath)
                    # 打开指定excel的指定sheet
                    # print(count_sheet_list[sheet_number])
                    EXCEL_df = pd.DataFrame(pd.read_excel(EXCELpath, sheet_name=count_sheet_list[sheet_number]))
                    for col_item in col_to_add[sheet_number]:
                        # 列合并
                        # 由于原先dataframe中有很多NaN空元素，无法比较。所以使用fillna('')函数将这些空元素以空字符填充
                        # print('sheet_number:', sheet_number, 'name:', name, 'col_choose_list[col_item]:', col_choose_list[col_item])
                        for row in range(MODEL_df.shape[0]):
                            if EXCEL_df[col_choose_list[col_item]].fillna('')[row] != MODEL_df[col_choose_list[col_item]].fillna('')[row]:
                                #MODEL_df[col_choose_list[col_item]][row] = MODEL_df[col_choose_list[col_item]].fillna('')[row] + EXCEL_df[col_choose_list[col_item]].fillna('')[row]
                                MODEL_df.loc[row:row, col_choose_list[col_item]] = MODEL_df[col_choose_list[col_item]].fillna('')[row] + EXCEL_df[col_choose_list[col_item]].fillna('')[row]

                        #MODEL_df[col_choose_list[col_item]] = MODEL_df[col_choose_list[col_item]].fillna('') + EXCEL_df[col_choose_list[col_item]].fillna('')
                    # print(MODEL_df['合作平台商'+'\n'+'（如与银行签约）'])
        # 将处理好的dataframe转化为输出excel文件的某一个sheet
        MODEL_df.to_excel(writer, sheet_name=count_sheet_list[sheet_number])

    # 保存和关闭
    writer.save()
    writer.close()

def decode_str(s):
    value, charset = decode_header(s)[0]
    if charset:
        value = value.decode(charset)
    return value

def recv_email_by_pop3(email_address, email_password, pop_server_host, pop_server_port, match, number, path):

    count = 0
    try:
        # 连接pop服务器。如果没有使用SSL，将POP3_SSL()改成POP3()即可其他都不需要做改动
        email_server = poplib.POP3_SSL(host=pop_server_host, port=pop_server_port, timeout=10)
        print("[pop3]----connect server success, now will check username")
    except:
        return "[pop3]----sorry the given email server address connect time out"
    try:
        # 验证邮箱是否存在
        email_server.user(email_address)
        print("[pop3]----username exist, now will check password")
    except:
        return "[pop3]----sorry the given email address seem do not exist"
    try:
        # 验证邮箱密码是否正确
        email_server.pass_(email_password)
        print("[pop3]----password correct,now will list email")
    except:
        return "[pop3]----sorry the given username seem do not correct"

    # 邮箱中其收到的邮件的数量
    email_count = len(email_server.list()[1])

    # list()返回所有邮件的编号:
    resp, mails, octets = email_server.list()
    # 遍历所有的邮件
    # 根据设定遍历的数目number
    if number == 0:
        # 如果设定nunber=0，则不进行限制
        circletime = len(mails) + 1
    elif number > (len(mails) + 1):
        # 如果设定nunber>邮件总数，则循环数为邮件总数
        circletime = len(mails) + 1
    else:
        # 如果设定nunber<=邮件总数，则循环数为number,因为是最后一种清况，所以用else就行了，不用elif
        circletime = number
    for i in range(1, circletime):
        # 通过retr(index)读取第index封邮件的内容；这里读取最后一封，也即最新收到的那一封邮件
        resp, lines, octets = email_server.retr(i)
        # lines是邮件内容，列表形式使用join拼成一个byte变量
        email_content = b'\r\n'.join(lines)
        try:
            # 再将邮件内容由byte转成str类型
            email_content = email_content.decode('utf-8')
        except Exception as e:
            print(str(e))
            continue
        # # 将str类型转换成<class 'email.message.Message'>
        # msg = email.message_from_string(email_content)
        msg = Parser().parsestr(email_content)
        # msg.get()能获取'To'收件人，'From'发件人，'Subject'邮件主题，这里我们只需要邮件主题。
        mailname = msg.get('Subject', '')
        print("[Mailname]:", decode_str(mailname))
        # 筛选邮件
        # 正则化匹配，如果邮件名decode_str(mailname)中包含match'走访情况记录'，则符合要求
        select = re.search(match, decode_str(mailname))
        print("--Condition select:", select)
        print("--count:", count)
        if select != None:
            # 获取附件
            count = count + 1
            f_list = get_att(msg, count, path)
            # print("f_list",f_list)

    # 关闭连接
    email_server.close()
    return 'success'

def get_att(msg, count, path):
    import email
    # 初始化附件名序列
    attachment_files = []

    for part in msg.walk():
        # 获取附件名称类型
        file_name = part.get_filename()
        contType = part.get_content_type()

        # 如果有附件
        if file_name:
            # 对附件名称进行解码
            dh = email.header.decode_header(file_name)
            # 举例：dh: [(b'=?UTF-8?Q?=E8=B5=B0=E8=AE=BF=E6=83=85=E5=86=B5=E8=AE=B0=E5=BD=950305.xlsx?=', 'us-ascii')]
            # dh[0][0]是邮件名 dh[0][1]是编码类型
            filename = dh[0][0]
            if dh[0][1]:
                # 根据编码类型dh[0][1]将附件名称可读化
                filename = decode_str(str(filename, dh[0][1]))
                print("--Attachment filename:", filename)
                # 重新设置邮件名
                filename_count_path = str(count) + '.xlsx'
            # 下载附件
            data = part.get_payload(decode=True)
            # 在指定目录下创建文件，注意二进制文件需要用wb模式打开
            att_file = open(path + '\\' + filename_count_path, 'wb')
            print(path + filename_count_path)
            # 在attachment_files这个序列后，添加本次循环内的附件名
            attachment_files.append(filename_count_path)
            # 保存附件
            att_file.write(data)
            att_file.close()
    return attachment_files


if __name__ == "__main__":#email_address, email_password, pop_server_host, pop_server_port
    global userjson
    userjson = {
        'email_address' : '',
        'email_password' : '',
        'pop_server_host' : ''
    }
    file_all = []
    for root, dirs, files in os.walk(os.getcwd()):
        file_all.extend(files)
    #print(file_all)
    if 'userinformation.json' in file_all:
        with open('userinformation.json', 'r') as f:
            userjson = json.load(f)
    else:
        with open('userinformation.json', 'w') as f:
            json.dump(userjson, f)

    #print(userjson)
    app = wx.App()
    frame0 = Main(None)
    frame0.Show()
    app.MainLoop()