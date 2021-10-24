wxIMPLEMENT_APP(MyApp);

bool MyApp::OnInit()
{
    if (!wxApp::OnInit())
        return false;

    MyFrame* frame = new MyFrame(L"HR Helper v." + f_GetWString_Version());
    frame->Show(true);


    Frame_Main = frame;
    g_hWnd = frame->GetHandle();


    Frame_Main->SetSize(g_FrameWidth, g_FrameHeight);

    //frame->m_timer.Start(1000);


    CSimpleIniA ini;
    ini.SetUnicode();

    if (f_Does_file_exist(g_INIPath) == 0)
    {
        ini.SetValue("Settings", "iLastVersion", "0");
        ini.SetValue("Settings", "bWriteLog", "0");

        ini.SetValue("Settings", "bLoadBaseAutomatically", "1");
        ini.SetValue("Settings", "bLoadBaseFile", "No_File");

        ini.SetValue("Settings", "bLoadReestrOrders", "1");
        ini.SetValue("Settings", "bLoadReestrOrdersFile", "No_File");

        ini.SetValue("Settings", "bLoadReestrDocs", "1");
        ini.SetValue("Settings", "bLoadReestrDocsFile", "No_File");

        ini.SetValue("Settings", "bLoadReestrAgreements", "1");
        ini.SetValue("Settings", "bLoadReestrAgreementsFile", "No_File");


        ini.SetValue("Settings", "bLoadBackupFolder", "1");
        ini.SetValue("Settings", "bLoadBackupFolderAlt", "1");
        ini.SetValue("Settings", "bLoadBackupFolderPath", "No_File");
        ini.SetValue("Settings", "bLoadBackupFolderPathAlt", "No_File");

        ini.SetValue("BackupFiles", "sBackupFilePath", "No_File");


        ini.SetValue("Settings", "bLoadMailMergeBaseEE", "1");
        ini.SetValue("Settings", "bLoadMailMergeBaseEEFile", "No_File");

        ini.SetValue("Settings", "bLoadMailMergeAttachmentFolder", "1");
        ini.SetValue("Settings", "bLoadMailMergeAttachmentFolderPath", "No_File");

        ini.SetLongValue("Settings", "iLastPanel", -1);

        ini.SetLongValue("Settings", "bMailMerge_LookForAttachments", 1);
        ini.SetLongValue("Settings", "bMailMerge_ShowLetters", 0);
        ini.SetLongValue("Settings", "b_Backup_OpenFolder", 1);
        ini.SetLongValue("Settings", "bMailMerge_LookOnlyByAttachments", 0);
        ini.SetLongValue("Settings", "bMailMerge_AddStaticAttachments", 1);
        ini.SetLongValue("Settings", "bMailMerge_SendOnlyOneLetter", 0);



        ini.SetLongValue("Settings", "b_DebugMode", 1);




        ini.SaveFile(g_INIPath.c_str());
    }

    auto errVal = ini.LoadFile(g_INIPath.c_str());

    if (errVal != SI_Error::SI_FILE)
    {

        string s_FilePath{};
        wstring w_FilePath{};


        if (ini.GetLongValue("Settings", "bLoadReestrOrders") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadReestrOrdersFile");
            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_Reestr_Orders, w_FilePath, 1);
                }
            }
        }

        if (ini.GetLongValue("Settings", "bLoadReestrDocs") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadReestrDocsFile");
            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_Reestr_Docs, w_FilePath, 1);
                }
            }
        }

        if (ini.GetLongValue("Settings", "bLoadReestrAgreements") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadReestrAgreementsFile");
            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_Reestr_Agreements, w_FilePath, 1);
                }
            }
        }



        if (ini.GetLongValue("Settings", "bLoadBackupFolder") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadBackupFolderPath");

            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_Backup_Folder, w_FilePath, 1);
                }
            }

        }

        if (ini.GetLongValue("Settings", "bLoadBackupFolderAlt") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadBackupFolderPathAlt");

            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_Backup_FolderAlt, w_FilePath, 1);
                }
            }

        }




        s_FilePath = ini.GetValue("BackupFiles", "sBackupFilePath");

        if (s_FilePath != "No_File")
        {


            std::vector<std::string> s_AllPaths = SplitString(s_FilePath, "|");
            for (std::vector<string>::iterator it = s_AllPaths.begin(); it != s_AllPaths.end(); ++it)
            {
                if ((*it).length() > 0)
                {
                    f_Open_File(t_File_Backup_Files, f_string_to_wString((*it)), 1);
                }

            }

        }


        if (ini.GetLongValue("Settings", "bLoadMailMergeBaseEE") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadMailMergeBaseEEFile");

            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_MailMerge_EEBase, w_FilePath, 1);
                }
            }

        }


        if (ini.GetLongValue("Settings", "bLoadMailMergeAttachmentFolder") == 1)
        {
            s_FilePath = ini.GetValue("Settings", "bLoadMailMergeAttachmentFolderPath");

            if (s_FilePath.length() > 0)
            {
                w_FilePath = f_string_to_wString(s_FilePath);

                if (f_Does_file_exist(w_FilePath) == 1)
                {
                    f_Open_File(t_File_MailMerge_AttachmentFolder, w_FilePath, 1);
                }
            }

        }

        int iLastPanel = ini.GetLongValue("Settings", "iLastPanel");


        g_MailMerge_SendOnlyOneLetter = ini.GetLongValue("Settings", "bMailMerge_SendOnlyOneLetter");
        g_MailMerge_LookforAttachments = ini.GetLongValue("Settings", "bMailMerge_LookForAttachments");
        g_MailMerge_ShowLetters = ini.GetLongValue("Settings", "bMailMerge_ShowLetters");
        g_DebugMode = ini.GetLongValue("Settings", "b_DebugMode");
        g_MailMerge_LookOnlyByAttachments = ini.GetLongValue("Settings", "bMailMerge_LookOnlyByAttachments");
        g_Backup_OpenFolder = ini.GetLongValue("Settings", "b_Backup_OpenFolder");
        g_MailMerge_AddStaticAttachments = ini.GetLongValue("Settings", "bMailMerge_AddStaticAttachments");

        g_UseReestrSearch = ini.GetLongValue("Settings", "bUseReestrSearch");


        if (g_UseReestrSearch != 1)
        {
            Frame_Main->m_switch_Button_Reestr->Show(0);
            if (iLastPanel == t_Panel_Reestr)
            {
                iLastPanel = t_Panel_MailMerge;
            }

        }




        g_UseVacSchedule = ini.GetLongValue("Settings", "bUseVacSchedule");


        if (g_UseVacSchedule != 1)
        {
            Frame_Main->m_switch_Button_VacSchedule->Show(0);
            if (iLastPanel == t_Panel_VacSchedule)
            {
                iLastPanel = t_Panel_MailMerge;
            }

        }



        Frame_Main->m_MailMerge_Checkbox_SendOnlyOneLetter->SetValue(g_MailMerge_SendOnlyOneLetter);
        Frame_Main->m_MailMerge_Checkbox_LookforAttachments->SetValue(g_MailMerge_LookforAttachments);
        Frame_Main->m_MailMerge_Checkbox_ShowLetters->SetValue(g_MailMerge_ShowLetters);
        Frame_Main->m_MailMerge_Checkbox_AddStaticAttachments->SetValue(g_MailMerge_AddStaticAttachments);
        Frame_Main->m_Backup_Checkbox_OpenFolder->SetValue(g_Backup_OpenFolder);
        Frame_Main->m_Checkbox_DebugMode->SetValue(g_DebugMode);


        f_Add_Menus();




        if (iLastPanel > -1 && iLastPanel < t_Panel_Max)
        {
            f_Panel_HideAllExcept(iLastPanel);
        }
        else {
            f_Panel_HideAllExcept(t_Panel_MailMerge);
        }

    }




    return true;

}