void MyFrame::OnBackupButtonPressed(wxCommandEvent& WXUNUSED(event))
{

    int iFolderExists = 0;
    int iFolderAltExists = 0;


    f_SetBusy(1);

    m_text->AppendText("Backing UP files\n");

    wxArrayString items;
    Frame_Main->m_Backup_ListBoxFiles->GetStrings(items);
    int num = items.GetCount();

    if (num <= 0)
    {
        m_text->AppendText("No files specified for backup\n");
        f_SetBusy(0);
        return;
    }



    if (f_Does_file_exist(g_BackupFolderPath))
    {
        m_text->AppendText("Backup folder exists\n");
        iFolderExists = 1;
    }

    if (f_Does_file_exist(g_BackupFolderPathAlt))
    {
        m_text->AppendText("Backup Alt folder exists\n");

        iFolderAltExists = 1;
    }

    if (iFolderExists == 0 && iFolderAltExists == 0)
    {
        m_text->AppendText("No folders exist\n");
        f_SetBusy(0);
        return;
    }


    if (iFolderExists)
    {
        int iFolderCreated = 0;
        wstring w_FolderName{};

            string s_FolderName{};
            wstring wp{};
            wstring w_tempString{};

            for (int i = 0; i < num; i++)
            {
                w_tempString = items[i];

                if (w_tempString.length() > 0 && f_Does_file_exist(w_tempString))
                {

                    if (iFolderCreated == 0)
                    {
                        iFolderCreated = 1;

                        time_t CurrentTime_T = time(0);

                        struct tm* Current_Time_TM = gmtime(&CurrentTime_T);

                        w_FolderName = g_BackupFolderPath + L"\\Backup_" + to_wstring(Current_Time_TM->tm_year + 1900) + "_";


                        int iMonth = Current_Time_TM->tm_mon + 1;
                        if (iMonth < 10)
                            w_FolderName += L"0";

                        w_FolderName += to_wstring(iMonth) + "_";

                        int iDay = Current_Time_TM->tm_mday;
                        if (iDay < 10)
                            w_FolderName += L"0";

                        w_FolderName += to_wstring(iDay) + "_";

                        int iHour = Current_Time_TM->tm_hour + 3;
                        if (iHour < 10)
                            w_FolderName += L"0";
                        w_FolderName += to_wstring(iHour) + "_";

                        int iMinute = Current_Time_TM->tm_min;
                        if (iMinute < 10)
                            w_FolderName += L"0";
                        w_FolderName += to_wstring(iMinute) + "_";

                        int iSecond = Current_Time_TM->tm_sec;
                        if (iSecond < 10)
                            w_FolderName += L"0";
                        w_FolderName += to_wstring(iSecond);


                        if (!f_Does_file_exist(w_FolderName))
                            std::filesystem::create_directory(w_FolderName);


                    }

                    std::filesystem::copy(w_tempString, w_FolderName);


                }

            }
        
        m_text->AppendText("Backup succesfull\n");

        if (g_Backup_OpenFolder && f_Does_file_exist(w_FolderName))
            ShellExecute(NULL, _T("explore"), w_FolderName.c_str(), NULL, NULL, SW_SHOW);
    }

    const auto copyOptions = std::filesystem::copy_options::overwrite_existing;

    if (iFolderAltExists)
    {
        wstring w_tempString{};
        for (int i = 0; i < num; i++)
        {
            w_tempString = items[i];

            if (w_tempString.length() > 0 && f_Does_file_exist(w_tempString))
            {
                std::filesystem::copy(w_tempString, g_BackupFolderPathAlt, copyOptions);
            }

        }

        m_text->AppendText("Backup Alt succesfull\n");

        if (g_Backup_OpenFolder && f_Does_file_exist(g_BackupFolderPathAlt))
            ShellExecute(NULL, _T("explore"), g_BackupFolderPathAlt.c_str(), NULL, NULL, SW_SHOW);
    }


  
    f_SetBusy(0);

}


void MyFrame::OnBackupButtonAddFolder(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening folder choose dialog for backup\n");
    f_Open_File_Dialog(t_File_Backup_Folder);
}

void MyFrame::OnBackupButtonAddFolderAlt(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening ALT folder choose dialog for backup\n");
    f_Open_File_Dialog(t_File_Backup_FolderAlt);
}

void MyFrame::OnBackupButtonAddFiles(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening files choose dialog for backup\n");
    f_Open_File_Dialog(t_File_Backup_Files);
}

void MyFrame::OnBackupButtonSaveFilesToINI(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Saving files to INI\n");

    string s_ToSave{};
    wstring w_ToSave{};
    wxArrayString items;
    Frame_Main->m_Backup_ListBoxFiles->GetStrings(items);

    int num = items.GetCount();

    CSimpleIniA ini;
    ini.SetUnicode();
    auto errVal = ini.LoadFile(g_INIPath.c_str());


    if (num > 0)
    {

        wxString str{};

        for (int i = 0; i < num; i++)
        {
            str = items[i];

            if (str.length() > 0)
            {
                w_ToSave += str;

                if (i != num - 1)
                    w_ToSave += L"|";
            }
        }

        if (errVal != SI_Error::SI_FILE)
        {
            s_ToSave = f_wString_to_string(w_ToSave);
            ini.SetValue("BackupFiles", "sBackupFilePath", s_ToSave.c_str());
            ini.SaveFile(g_INIPath.c_str());
        }

    }
    else {

        if (errVal != SI_Error::SI_FILE)
        {
            ini.SetValue("BackupFiles", "sBackupFilePath", "No_File");
            ini.SaveFile(g_INIPath.c_str());
        }
    }

}