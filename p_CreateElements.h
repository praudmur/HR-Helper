


MyFrame::MyFrame(const wxString& title)
    : wxFrame(NULL, wxID_ANY, title)
{




    SetIcon(wxICON(sample));




    wxPoint iPoint{};
    wxPoint iPoint2{};
    wxSize iSize{};

    wxPanel* panel = new wxPanel(this);
    g_panel_Reestr = panel;

    panel->Bind(wxEVT_CHAR_HOOK, &MyFrame::OnKeyPress, this); //////////////////////EVENTS
    panel->Bind(wxEVT_TEXT_ENTER, &MyFrame::OnReestrButtonPressed, this); //////////////////////EVENTS



    //panel->Bind(wxEVT_CLOSE_WINDOW, &MyFrame::OnQuit, this);
    //Connect(wxEVT_CLOSE_WINDOW, wxCloseEventHandler(Veto::OnQuit));



    //wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    //panel->SetSizer(sizer);

    Panel PanelReestr;
    PanelReestr.iID = t_Panel_Reestr;
    m_switch_Button_Reestr = new wxButton(panel, wx_SwitchPanelsReestr, "�������� ��������", wxPoint(10, 10));
    PanelReestr.SwitchButton = m_switch_Button_Reestr;

    Panel PanelMailMerge;
    PanelMailMerge.iID = t_Panel_MailMerge;
    m_switch_Button_MailMerge = new wxButton(panel, wx_SwitchPanelsMailMerge, "��������", wxPoint(230, 10));
    PanelMailMerge.SwitchButton = m_switch_Button_MailMerge;
    PanelMailMerge.iShowDebugText = 0;


    Panel PanelDateDifference;
    PanelDateDifference.iID = t_Panel_DateDifference;
    m_switch_Button_DateDiff = new wxButton(panel, wx_SwitchPanelsDateDifference, "������� ���", wxPoint(600, 10));
    PanelDateDifference.SwitchButton = m_switch_Button_DateDiff;

    Panel PanelBackup;
    PanelBackup.iID = t_Panel_Backup;
    m_switch_Button_Backup = new wxButton(panel, wx_SwitchPanelsBackup, "��������� �����", wxPoint(400, 10));
    PanelBackup.SwitchButton = m_switch_Button_Backup;





    m_Checkbox_DebugMode = new wxCheckBox(panel, wx_chkBox_All, "Debug mode", wxPoint(1100, 10));


    m_reestr_Button_Search = new wxButton(panel, wx_reestr_StartSearchingReestrs, "������ �����", wxPoint(300, 70));
    m_dateDiff_Button_Execute = new wxButton(panel, wx_DateDiff_ExecuteDateDifference, "���������� ������� � �����", wxPoint(300, 70));
    m_Backup_Button_Execute = new wxButton(panel, wx_Backup_ExecuteBackup, "������� ��������� �����", wxPoint(300, 70));
    m_MailMerge_Button_Execute = new wxButton(panel, wx_MailMerge_ExecuteMailMerge, "������ ��������", wxPoint(300, 70));


    m_text = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(0, 700), wxSize(g_FrameWidth, 200), wxTE_MULTILINE | wxTE_READONLY);


    //sizer->Add(m_text, 0, wxEXPAND);

    m_text->AppendText("����� ���������\n");

    iSize.x = 900;
    iSize.y = 80;


    m_reestr_Status_ReestrOrdersPath = new wxTextCtrl(panel, wx_reestr_OpenReestrOrders, "", wxPoint(415, 130), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrOrdersPath->SetLabelText("������ �������� �� ��������");
    m_reestr_Status_ReestrOrdersPath->SetBackgroundColour(*wxYELLOW);

    m_reestr_Status_ReestrDocsPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(415, 210), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrDocsPath->SetLabelText("������ ��������� ���������� �� ��������");
    m_reestr_Status_ReestrDocsPath->SetBackgroundColour(*wxYELLOW);

    m_reestr_Status_ReestrAgreementsPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(415, 290), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrAgreementsPath->SetLabelText("������ ���. ���������� �� ��������");
    m_reestr_Status_ReestrAgreementsPath->SetBackgroundColour(*wxYELLOW);


    iPoint.x = 415;
    iPoint.y = 370;
    iSize.x = 400;
    iSize.y = 60;

    //m_reestr_Search_EmployeeName = new wxTextCtrl(panel, wx_reestr_Search_EmployeeName, "������� ������� ���������", iPoint, iSize, wxTE_PROCESS_ENTER);

    m_reestr_Button_Mass_Reminder = new wxButton(panel, wx_reestr_StartMassReminder, "������ �������� ��������", wxPoint(500, 400));








    //DATE DIFFERENCE PANEL//////////////////////////////////////////


    iSize.x = 150;
    iSize.y = 30;
    iPoint.x = 30;
    iPoint.y = 130;


    m_dateDiff_DatePickerFrom = new wxDatePickerCtrl(panel, 6003, wxDefaultDateTime, iPoint, iSize, 2);

    iPoint.x = 190;
    iPoint.y = 130;

    m_dateDiff_DatePickerTo = new wxDatePickerCtrl(panel, 6003, wxDefaultDateTime, iPoint, iSize, 2);



    //Backup PANEL//////////////////////////////////////////


    m_Backup_ListBoxFiles = new wxEditableListBox(panel, wx_Backup_ListBox, ("������ ������:"), wxPoint(15, 120), wxSize(1000, 300), wxEL_DEFAULT_STYLE);


    m_Backup_Button_AddFiles = new wxButton(panel, wx_Backup_AddFiles, "�������� �����", wxPoint(15, 440));
    m_Backup_Button_SaveFilesToINI = new wxButton(panel, wx_Backup_SaveFilesToINI, "��������� ����� � INI", wxPoint(200, 440));




    m_Backup_Status_BackupFolderPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(15, 480), wxSize(1000, 60), wxTE_READONLY);
    //sizer->Add(m_text, wxSizerFlags(1).Expand().Border());
    m_Backup_Status_BackupFolderPath->SetLabelText("����� ��������� ����� �� �������");
    m_Backup_Status_BackupFolderPath->SetBackgroundColour(*wxYELLOW);


    m_Backup_Button_AddFolder = new wxButton(panel, wx_Backup_AddFolder, "������� �����", wxPoint(15, 600));

    m_Backup_Checkbox_OpenFolder = new wxCheckBox(panel, wx_chkBox_All, "������� ����� ����� �����������", wxPoint(700, 70));

    //MailMerge panel

    m_MailMerge_Text_LetterSubject = new wxTextCtrl(panel, wx_MailMerge_TextLetterBody, "", wxPoint(450, 130), wxSize(900, 60), wxTE_MULTILINE);
    m_MailMerge_Text_LetterSubject->AppendText("������� ���� ������");

    m_MailMerge_Text_LetterBody = new wxTextCtrl(panel, wx_MailMerge_TextLetterBody, "", wxPoint(450, 200), wxSize(900, 340), wxTE_MULTILINE);
    m_MailMerge_Text_LetterBody->AppendText("<���>, ������������!");

    m_MailMerge_Status_AttachmentFolderPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(450, 550), wxSize(900, 80), wxTE_READONLY | wxTE_MULTILINE);
    m_MailMerge_Status_AttachmentFolderPath->SetLabelText("����� �������� �� �������");
    m_MailMerge_Status_AttachmentFolderPath->SetBackgroundColour(*wxYELLOW);

    m_MailMerge_Button_AddAttachmentFolder = new wxButton(panel, wx_MailMerge_AddAttachmentFolder, "������� ����� ��������", wxPoint(450, 630));

    m_MailMerge_Checkbox_LookforAttachments = new wxCheckBox(panel, wx_chkBox_All, "������ ��������", wxPoint(550, 100));
    m_MailMerge_Checkbox_ShowLetters = new wxCheckBox(panel, wx_chkBox_All, "�������� ������", wxPoint(800, 100));
    m_MailMerge_Checkbox_AddStaticAttachments = new wxCheckBox(panel, wx_chkBox_All, "�������� ��������� �����", wxPoint(1050, 100));
    m_MailMerge_Checkbox_LookforAttachments->SetToolTip("������ �������� �� ��� ��������� � ����� ��������");
    m_MailMerge_Checkbox_ShowLetters->SetToolTip("������� ������ ����� ������������");
    m_MailMerge_Checkbox_AddStaticAttachments->SetToolTip("�������� ��������� � ������ �������� ����� �� ��� ������");

    m_MailMerge_ListBoxAttachments = new wxEditableListBox(panel, wx_MailMerge_ListBox, ("������ ��������:"), wxPoint(15, 680), wxSize(800, 200), wxEL_DEFAULT_STYLE);
    m_MailMerge_Button_AddAttachmentFiles = new wxButton(panel, wx_MailMerge_AddAttachmentFiles, "�������� ��������", wxPoint(15, 640));



    ////////////////////////////////////////////////

    //wxArrayString wx_Items{};
    //m_EEBaseList = new wxListBox(panel, wx_EEBase_ListSelected, wxPoint(15, 130), wxSize(400, 410), wx_Items, wxLB_SINGLE);



    ////////////////////////////////////////////
    m_EEList = new wxCheckListBox(panel, wxID_ANY, wxPoint(15, 130), wxSize(400, 410));
    m_EEList_Button_AddBase = new wxButton(panel, wx_EElist_AddEEBase, "�������� ����", wxPoint(15, 550));


    wxSize ButtonSize = m_EEList_Button_AddBase->GetSize();


    m_EEList_Button_SelectAll = new wxButton(panel, wx_EElist_SelectAll, "[ALL]", wxPoint(170, 550), wxSize(100, ButtonSize.GetHeight()));
    m_EEList_Button_SelectNone = new wxButton(panel, wx_EElist_SelectNone, "[NONE]", wxPoint(270, 550), wxSize(100, ButtonSize.GetHeight()));
    m_EEList_Button_SelectAll->SetToolTip("�������� ����");
    m_EEList_Button_SelectNone->SetToolTip("����� ��������� � ����������");


    g_DatePickersArray.push_back(m_dateDiff_DatePickerFrom);
    g_DatePickersArray.push_back(m_dateDiff_DatePickerTo);

    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrOrdersPath);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrDocsPath);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrAgreementsPath);
    //PanelReestr.Elements.push_back((wxControl*)m_reestr_Search_EmployeeName);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Button_Search);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Button_Mass_Reminder);
    

    PanelReestr.Elements.push_back((wxControl*)m_EEList);
    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_AddBase);
    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_SelectAll);
    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_SelectNone);




    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_Execute);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_AddBase);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Text_LetterBody);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Text_LetterSubject);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Status_AttachmentFolderPath);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_AddAttachmentFolder);

    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_LookforAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_ShowLetters);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_AddStaticAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_ListBoxAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_AddAttachmentFiles);

    
    

    PanelMailMerge.Elements.push_back((wxControl*)m_EEList);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_SelectAll);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_SelectNone);




    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_DatePickerFrom);
    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_DatePickerTo);
    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_Button_Execute);

    PanelBackup.Elements.push_back((wxControl*)m_Backup_ListBoxFiles);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_Execute);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_AddFolder);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_AddFiles);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Status_BackupFolderPath);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_SaveFilesToINI);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Checkbox_OpenFolder);




   


    g_PanelsArray.push_back(PanelReestr);
    g_PanelsArray.push_back(PanelDateDifference);
    g_PanelsArray.push_back(PanelBackup);
    g_PanelsArray.push_back(PanelMailMerge);



    //t_Panel_Max = g_PanelsArray.size();

    //for (std::vector<Panel>::iterator it = g_PanelsArray.begin(); it != g_PanelsArray.end(); ++it)
    //{
    //    m_text->AppendText(to_string((*it).iID));
    //}



}