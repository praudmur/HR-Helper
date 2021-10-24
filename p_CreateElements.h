



MyFrame::MyFrame(const wxString& title)
    : wxFrame(NULL, wxID_ANY, title)
{

   SetIcon(wxICON(aaaaappicon));

    wxPoint iPoint{};
    wxPoint iPoint2{};
    wxSize iSize{};

    wxScrolledWindow* panel = new wxScrolledWindow(this);
    g_panel_Reestr = panel;

    panel->SetScrollbars(20, 25, 50, 50);


    //panel->Show(0);
   // MyFrame *panel = this;

    panel->Bind(wxEVT_CHAR_HOOK, &MyFrame::OnKeyPress, this); //////////////////////EVENTS
    panel->Bind(wxEVT_TEXT_ENTER, &MyFrame::OnReestrButtonPressed, this); //////////////////////EVENTS


    //panel->Bind(wxEVT_CLOSE_WINDOW, &MyFrame::OnQuit, this);
    //Connect(wxEVT_CLOSE_WINDOW, wxCloseEventHandler(Veto::OnQuit));



    //wxBoxSizer* sizer = new wxBoxSizer(wxVERTICAL);
    //panel->SetSizer(sizer);

    Panel PanelReestr;
    PanelReestr.iID = t_Panel_Reestr;
    m_switch_Button_Reestr = new wxButton(panel, wx_SwitchPanelsReestr, "Проверка реестров", wxPoint(10, 10));
    PanelReestr.SwitchButton = m_switch_Button_Reestr;

    Panel PanelMailMerge;
    PanelMailMerge.iID = t_Panel_MailMerge;
    m_switch_Button_MailMerge = new wxButton(panel, wx_SwitchPanelsMailMerge, "Рассылки", wxPoint(230, 10));
    PanelMailMerge.SwitchButton = m_switch_Button_MailMerge;
    PanelMailMerge.iDebugTextPos = 950;
    PanelMailMerge.iScrollBarHeight = 24;

    //PanelMailMerge.iShowDebugText = 0;


    Panel PanelDateDifference;
    PanelDateDifference.iID = t_Panel_DateDifference;
    m_switch_Button_DateDiff = new wxButton(panel, wx_SwitchPanelsDateDifference, "Разница дат", wxPoint(600, 10));
    PanelDateDifference.SwitchButton = m_switch_Button_DateDiff;

    Panel PanelBackup;
    PanelBackup.iID = t_Panel_Backup;
    m_switch_Button_Backup = new wxButton(panel, wx_SwitchPanelsBackup, "Резервные копии", wxPoint(400, 10));
    PanelBackup.SwitchButton = m_switch_Button_Backup;


    Panel PanelVacSchedule;
    PanelVacSchedule.iID = t_Panel_VacSchedule;
    m_switch_Button_VacSchedule = new wxButton(panel, wx_SwitchPanelsVacSchedule, "График отпусков", wxPoint(750, 10));
    PanelVacSchedule.SwitchButton = m_switch_Button_VacSchedule;





    m_Checkbox_DebugMode = new wxCheckBox(panel, wx_chkBox_All, "Debug mode", wxPoint(1100, 10));


    m_reestr_Button_Search = new wxButton(panel, wx_reestr_StartSearchingReestrs, "Начать поиск", wxPoint(300, 70));
    m_dateDiff_Button_Execute = new wxButton(panel, wx_DateDiff_ExecuteDateDifference, "Рассчитать разницу в датах", wxPoint(300, 70));
    m_Backup_Button_Execute = new wxButton(panel, wx_Backup_ExecuteBackup, "Сделать резервную копию", wxPoint(300, 70));
    m_MailMerge_Button_Execute = new wxButton(panel, wx_MailMerge_ExecuteMailMerge, "Начать рассылку", wxPoint(300, 70));


    m_text = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(0, 700), wxSize(g_FrameWidthMinusScrollBar, 200), wxTE_MULTILINE | wxTE_READONLY);

    

    //sizer->Add(m_text, 0, wxEXPAND);

    m_text->AppendText("Старт программы\n");

    iSize.x = 900;
    iSize.y = 80;


    m_reestr_Status_ReestrOrdersPath = new wxTextCtrl(panel, wx_reestr_OpenReestrOrders, "", wxPoint(415, 130), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrOrdersPath->SetLabelText("Реестр приказов не загружен");
    m_reestr_Status_ReestrOrdersPath->SetBackgroundColour(*wxYELLOW);

    m_reestr_Status_ReestrDocsPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(415, 210), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrDocsPath->SetLabelText("Реестр получения документов не загружен");
    m_reestr_Status_ReestrDocsPath->SetBackgroundColour(*wxYELLOW);

    m_reestr_Status_ReestrAgreementsPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(415, 290), iSize, wxTE_READONLY | wxTE_MULTILINE);
    m_reestr_Status_ReestrAgreementsPath->SetLabelText("Реестр доп. соглашений не загружен");
    m_reestr_Status_ReestrAgreementsPath->SetBackgroundColour(*wxYELLOW);


    iPoint.x = 415;
    iPoint.y = 370;
    iSize.x = 400;
    iSize.y = 60;

    //m_reestr_Search_EmployeeName = new wxTextCtrl(panel, wx_reestr_Search_EmployeeName, "Введите Фамилию работника", iPoint, iSize, wxTE_PROCESS_ENTER);

    m_reestr_Button_Mass_Reminder = new wxButton(panel, wx_reestr_StartMassReminder, "Начать массовую рассылку", wxPoint(500, 400));








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


    m_Backup_ListBoxFiles = new wxEditableListBox(panel, wx_Backup_ListBox, ("Список файлов:"), wxPoint(15, 120), wxSize(1000, 300), wxEL_DEFAULT_STYLE);


    m_Backup_Button_AddFiles = new wxButton(panel, wx_Backup_AddFiles, "Добавить файлы", wxPoint(15, 440));
    m_Backup_Button_SaveFilesToINI = new wxButton(panel, wx_Backup_SaveFilesToINI, "Сохранить файлы в INI", wxPoint(200, 440));




    m_Backup_Status_BackupFolderPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(15, 490), wxSize(1000, 60), wxTE_READONLY);
    m_Backup_Status_BackupFolderPath->SetLabelText("Папка резервной копии не указана");
    m_Backup_Status_BackupFolderPath->SetBackgroundColour(*wxYELLOW);
    m_Backup_Status_BackupFolderPath->Bind(wxEVT_LEFT_DCLICK, &MyFrame::OnMouseDoubleClick, this);


    m_Backup_Status_BackupFolderPathAlt = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(15, 560), wxSize(1000, 60), wxTE_READONLY);
    m_Backup_Status_BackupFolderPathAlt->SetLabelText("Прямая папка резервной копии не указана");
    m_Backup_Status_BackupFolderPathAlt->SetBackgroundColour(*wxYELLOW);
    m_Backup_Status_BackupFolderPathAlt->Bind(wxEVT_LEFT_DCLICK, &MyFrame::OnMouseDoubleClick, this);


    m_Backup_Button_AddFolder = new wxButton(panel, wx_Backup_AddFolder, "Указать папку", wxPoint(1020, 490));
    m_Backup_Button_AddFolderAlt = new wxButton(panel, wx_Backup_AddFolderAlt, "Указать папку", wxPoint(1020, 560));


    m_Backup_Checkbox_OpenFolder = new wxCheckBox(panel, wx_chkBox_All, "Открыть папку после копирования", wxPoint(700, 70));

    //MailMerge panel

    m_MailMerge_Text_LetterSubject = new wxTextCtrl(panel, wx_MailMerge_TextLetterBody, "", wxPoint(450, 130), wxSize(900, 60), wxTE_MULTILINE);
    m_MailMerge_Text_LetterSubject->AppendText("Введите тему письма");

    m_MailMerge_Text_LetterBody = new wxTextCtrl(panel, wx_MailMerge_TextLetterBody, "", wxPoint(450, 200), wxSize(900, 340), wxTE_MULTILINE);
    m_MailMerge_Text_LetterBody->AppendText("<имя>, здравствуйте!");

    m_MailMerge_Status_AttachmentFolderPath = new wxTextCtrl(panel, wxID_ANY, "", wxPoint(450, 550), wxSize(900, 80), wxTE_READONLY | wxTE_MULTILINE);
    m_MailMerge_Status_AttachmentFolderPath->SetLabelText("Папка вложений не указана");
    m_MailMerge_Status_AttachmentFolderPath->SetBackgroundColour(*wxYELLOW);
    m_MailMerge_Status_AttachmentFolderPath->Bind(wxEVT_LEFT_DCLICK, &MyFrame::OnMouseDoubleClick, this);



    m_MailMerge_Button_AddAttachmentFolder = new wxButton(panel, wx_MailMerge_AddAttachmentFolder, "Выбрать папку вложений", wxPoint(450, 630));


    m_MailMerge_Checkbox_SendOnlyOneLetter = new wxCheckBox(panel, wx_chkBox_All, "Отправить в одном письме", wxPoint(550, 60));
    m_MailMerge_Checkbox_LookforAttachments = new wxCheckBox(panel, wx_chkBox_All, "Искать вложения", wxPoint(550, 100));
    m_MailMerge_Checkbox_ShowLetters = new wxCheckBox(panel, wx_chkBox_All, "Показать письма", wxPoint(800, 100));
    m_MailMerge_Checkbox_AddStaticAttachments = new wxCheckBox(panel, wx_chkBox_All, "Добавить выбранные файлы", wxPoint(1030, 100));

    m_MailMerge_Checkbox_SendOnlyOneLetter->SetToolTip("Отправить текст всем адресатам в одном письме");
    m_MailMerge_Checkbox_LookforAttachments->SetToolTip("Искать вложения по ФИО работника в папке вложений");
    m_MailMerge_Checkbox_ShowLetters->SetToolTip("Открыть письма после формирования");
    m_MailMerge_Checkbox_AddStaticAttachments->SetToolTip("Добавить указанные в Списке вложений файлы во все письма");

    m_MailMerge_ListBoxAttachments = new wxEditableListBox(panel, wx_MailMerge_ListBox, ("Список вложений:"), wxPoint(450, 680), wxSize(900, 200), wxEL_DEFAULT_STYLE);
    m_MailMerge_Button_AddAttachmentFiles = new wxButton(panel, wx_MailMerge_AddAttachmentFiles, "Добавить вложения", wxPoint(450, 890));



    ////////////////////////////////////////////////

    wxArrayString wx_Items{};
    m_VacSchedule_EEList = new wxListBox(panel, wx_VacSchedule_EEList, wxPoint(15, 60), wxSize(400, 410), wx_Items, wxLB_SINGLE);
    m_VacSchedule_Reestr = new wxListCtrl(panel,wx_VacSchedule_Reestr, wxPoint(420, 60), wxSize(g_FrameWidth - 460, 410), wxLC_REPORT);
    m_VacSchedule_EEVacList = new wxListCtrl(panel, wx_VacSchedule_EEVacList, wxPoint(420, 480), wxSize(g_FrameWidth - 460, 200), wxLC_REPORT);





    ////////////////////////////////////////////
    m_EEList = new wxCheckListBox(panel, wxID_ANY, wxPoint(15, 130), wxSize(400, 410));
    m_EEListDepartment = new wxCheckListBox(panel, wx_EElist_Department, wxPoint(15, 600), wxSize(400, 210));

    m_EEList_Button_AddBase = new wxButton(panel, wx_EElist_AddEEBase, "Добавить базу", wxPoint(15, 550));


    wxSize ButtonSize = m_EEList_Button_AddBase->GetSize();


    m_EEList_Button_SelectAll = new wxButton(panel, wx_EElist_SelectAll, "[ALL]", wxPoint(170, 550), wxSize(100, ButtonSize.GetHeight()));
    m_EEList_Button_SelectNone = new wxButton(panel, wx_EElist_SelectNone, "[NONE]", wxPoint(270, 550), wxSize(100, ButtonSize.GetHeight()));
    m_EEList_Button_SelectAll->SetToolTip("Выделить всех");
    m_EEList_Button_SelectNone->SetToolTip("Снять выделение с работников");


    g_DatePickersArray.push_back(m_dateDiff_DatePickerFrom);
    g_DatePickersArray.push_back(m_dateDiff_DatePickerTo);

    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrOrdersPath);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrDocsPath);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Status_ReestrAgreementsPath);
    //PanelReestr.Elements.push_back((wxControl*)m_reestr_Search_EmployeeName);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Button_Search);
    PanelReestr.Elements.push_back((wxControl*)m_reestr_Button_Mass_Reminder);
    
    PanelReestr.Elements.push_back((wxControl*)m_EEList);
    PanelReestr.Elements.push_back((wxControl*)m_EEListDepartment);

    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_AddBase);
    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_SelectAll);
    PanelReestr.Elements.push_back((wxControl*)m_EEList_Button_SelectNone);

    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_Execute);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_AddBase);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Text_LetterBody);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Text_LetterSubject);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Status_AttachmentFolderPath);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_AddAttachmentFolder);

    
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_SendOnlyOneLetter);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_LookforAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_ShowLetters);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Checkbox_AddStaticAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_ListBoxAttachments);
    PanelMailMerge.Elements.push_back((wxControl*)m_MailMerge_Button_AddAttachmentFiles);

    PanelMailMerge.Elements.push_back((wxControl*)m_EEList);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEListDepartment);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_SelectAll);
    PanelMailMerge.Elements.push_back((wxControl*)m_EEList_Button_SelectNone);

    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_DatePickerFrom);
    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_DatePickerTo);
    PanelDateDifference.Elements.push_back((wxControl*)m_dateDiff_Button_Execute);

    PanelBackup.Elements.push_back((wxControl*)m_Backup_ListBoxFiles);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_Execute);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_AddFolder);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_AddFolderAlt);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_AddFiles);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Status_BackupFolderPath);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Status_BackupFolderPathAlt);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Button_SaveFilesToINI);
    PanelBackup.Elements.push_back((wxControl*)m_Backup_Checkbox_OpenFolder);



    

    PanelVacSchedule.Elements.push_back((wxControl*)m_VacSchedule_EEList);
    PanelVacSchedule.Elements.push_back((wxControl*)m_VacSchedule_Reestr);
    PanelVacSchedule.Elements.push_back((wxControl*)m_VacSchedule_EEVacList);
    
    
   




    g_PanelsArray.push_back(PanelReestr);
    g_PanelsArray.push_back(PanelDateDifference);
    g_PanelsArray.push_back(PanelBackup);
    g_PanelsArray.push_back(PanelMailMerge);
    g_PanelsArray.push_back(PanelVacSchedule);



    t_Panel_Max = g_PanelsArray.size();


    for (std::vector<Panel>::iterator it = g_PanelsArray.begin(); it != g_PanelsArray.end(); ++it)
    {
        for (std::vector<wxControl*>::iterator it2 = Iter.Elements.begin(); it2 != Iter.Elements.end(); ++it2)
        {
            Iter2->Show(0);
        }
    }



}