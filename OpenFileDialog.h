

void f_Clear_File(int iRequest)
{
    if (iRequest == t_File_MailMerge_AttachmentFolder)
    {
        g_MailMergeAttachmentFolderPath = L"No_File";
        f_SetINIValue("Settings", "bLoadMailMergeAttachmentFolderPath", L"No_File");
        Frame_Main->m_MailMerge_Status_AttachmentFolderPath->SetLabelText("����� �������� �� �������");
        Frame_Main->m_MailMerge_Status_AttachmentFolderPath->SetBackgroundColour(*wxYELLOW);

    }
    else if (iRequest == t_File_Backup_Folder)
    {
        g_BackupFolderPath = L"No_File";
        f_SetINIValue("Settings", "bLoadBackupFolderPath", L"No_File");
        Frame_Main->m_Backup_Status_BackupFolderPath->SetLabelText("����� ��������� ����� �� �������");
        Frame_Main->m_Backup_Status_BackupFolderPath->SetBackgroundColour(*wxYELLOW);

    }
    else if (iRequest == t_File_Backup_FolderAlt)
    {
        g_BackupFolderPathAlt = L"No_File";
        f_SetINIValue("Settings", "bLoadBackupFolderPathAlt", L"No_File");
        Frame_Main->m_Backup_Status_BackupFolderPathAlt->SetLabelText("������ ����� ��������� ����� �� �������");
        Frame_Main->m_Backup_Status_BackupFolderPathAlt->SetBackgroundColour(*wxYELLOW);

    }


}



void f_Open_File(int iRequest, wstring w_FilePath, bool bFromINI)
{


    string s_Message{};
    string s_String_To_Store{};
    if (iRequest == t_File_Reestr_Docs)
    {
        if (g_ReestrDocsPathLoaded == 0)
        {
            g_FilesLoaded += 1;
        }

        g_ReestrDocsPathLoaded = 1;

        g_ReestrDocsPath = w_FilePath;

        Frame_Main->m_text->AppendText("\n���� ��������::");
        Frame_Main->m_text->AppendText(g_ReestrDocsPath);
        s_Message = "������ ��������>>>>>" + g_ReestrDocsPath;

        Frame_Main->m_reestr_Status_ReestrDocsPath->SetLabelText(s_Message);
        Frame_Main->m_reestr_Status_ReestrDocsPath->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadReestrDocsFile", w_FilePath);
        }
    }
    else if (iRequest == t_File_Reestr_Orders)
    {
        if (g_ReestrOrdersPathLoaded == 0)
        {
            g_FilesLoaded += 1;
        }

        g_ReestrOrdersPathLoaded = 1;

        g_ReestrOrdersPath = w_FilePath;

        Frame_Main->m_text->AppendText("\n���� ��������::");
        Frame_Main->m_text->AppendText(g_ReestrOrdersPath);
        s_Message = "������ ��������>>>>>" + g_ReestrOrdersPath;

        Frame_Main->m_reestr_Status_ReestrOrdersPath->SetLabelText(s_Message);
        Frame_Main->m_reestr_Status_ReestrOrdersPath->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadReestrOrdersFile", w_FilePath);
        }

    }
    else if (iRequest == t_File_Reestr_Agreements)
    {
        if (g_ReestrAgreementsPathLoaded == 0)
        {
            g_FilesLoaded += 1;
        }

        g_ReestrAgreementsPathLoaded = 1;

        g_ReestrAgreementsPath = w_FilePath;

        Frame_Main->m_text->AppendText("\n���� ��������::");
        Frame_Main->m_text->AppendText(g_ReestrAgreementsPath);
        s_Message = "������ ��������>>>>>" + g_ReestrAgreementsPath;

        Frame_Main->m_reestr_Status_ReestrAgreementsPath->SetLabelText(s_Message);
        Frame_Main->m_reestr_Status_ReestrAgreementsPath->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadReestrAgreementsFile", w_FilePath);
        }


    }else if (iRequest == t_File_Backup_Files)
    {

        //Frame_Main->m_text->AppendText("\nOPENING");
        wxArrayString items;
        Frame_Main->m_Backup_ListBoxFiles->GetStrings(items);
        items.Add(w_FilePath);


        Frame_Main->m_Backup_ListBoxFiles->SetStrings(items);



    }
    else if (iRequest == t_File_Backup_Folder)
    {

        g_BackupFolderPath = w_FilePath;
        s_Message = "����� ��������� �����>>>" + g_BackupFolderPath;

        Frame_Main->m_Backup_Status_BackupFolderPath->SetLabelText(s_Message);
        Frame_Main->m_Backup_Status_BackupFolderPath->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadBackupFolderPath", w_FilePath);
        }
    }else if (iRequest == t_File_Backup_FolderAlt)
    {

        g_BackupFolderPathAlt = w_FilePath;
        s_Message = "����� ������ ��������� �����>>>" + g_BackupFolderPathAlt;

        Frame_Main->m_Backup_Status_BackupFolderPathAlt->SetLabelText(s_Message);
        Frame_Main->m_Backup_Status_BackupFolderPathAlt->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadBackupFolderPathAlt", w_FilePath);
        }
    }else if (iRequest == t_File_MailMerge_EEBase)
    {

        g_MailMergeEEBasePath = w_FilePath;

        s_Message = "���� ���������� Mail Merge>>>" + g_MailMergeEEBasePath;
        Frame_Main->m_text->AppendText(s_Message);

        f_Excel_OpenEEBase();

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadMailMergeBaseEEFile", w_FilePath);
        }

    }else if (iRequest == t_File_MailMerge_AttachmentFolder)
    {


            g_MailMergeAttachmentFolderPath = w_FilePath;
            s_Message = "����� ��������>>>" + g_MailMergeAttachmentFolderPath;

            Frame_Main->m_MailMerge_Status_AttachmentFolderPath->SetLabelText(s_Message);
            Frame_Main->m_MailMerge_Status_AttachmentFolderPath->SetBackgroundColour(*wxGREEN);

            if (!bFromINI)
            {
                f_SetINIValue("Settings", "bLoadMailMergeAttachmentFolderPath", w_FilePath);
            }





    }else if (iRequest == t_File_MailMerge_AttachmentFiles)
    {
        wxArrayString items;
        Frame_Main->m_MailMerge_ListBoxAttachments->GetStrings(items);
        items.Add(w_FilePath);
        Frame_Main->m_MailMerge_ListBoxAttachments->SetStrings(items);


        //Frame_Main->m_MailMerge_ListBoxAttachments->GetListCtrl()->SetItemTextColour(0,*wxGREEN);

    }
    else if (iRequest == t_File_VacSchedule_EEBase)
    {

        g_VacScheduleEEBasePath = w_FilePath;
        Frame_Main->m_text->AppendText(L"���� ���������� Vac Schedule>>>" + g_VacScheduleEEBasePath);

        f_VacSchedule_OpenEEBase();

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadVacScheduleBaseEEFile", w_FilePath);
        }
    }

}



int WINAPI f_Open_File_Dialog(int iRequest)
{
    int iMultiSelect = 0;


    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED |
        COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        IFileOpenDialog* pFileOpen;



        hr = CoCreateInstance(CLSID_FileOpenDialog, NULL, CLSCTX_ALL,
            IID_IFileOpenDialog, reinterpret_cast<void**>(&pFileOpen));


        if (iRequest == t_File_Reestr_Docs)
        {
            pFileOpen->SetTitle(L"�������� ������ ��������� ����������");
        }
        else if (iRequest == t_File_Reestr_Orders)
        {
            pFileOpen->SetTitle(L"�������� ������ ��������");
        }else if (iRequest == t_File_Reestr_Agreements)
        {
            pFileOpen->SetTitle(L"�������� ������ ���. ����������");
        }
        else if (iRequest == t_File_Backup_Files)
        {
            pFileOpen->SetTitle(L"�������� �����, ������� ����� �����������");
        }else if (iRequest == t_File_Backup_Folder)
        {
            pFileOpen->SetTitle(L"�������� �����, ���� ����� ����������� �����");
        }
        else if (iRequest == t_File_Backup_FolderAlt)
        {
            pFileOpen->SetTitle(L"�������� �����, ���� ����� ����������� ����� - ������ �����");
        }
        else if (iRequest == t_File_MailMerge_EEBase)
        {
            pFileOpen->SetTitle(L"�������� ���� ���� ����������");
        }else if (iRequest == t_File_MailMerge_AttachmentFolder)
        {
            pFileOpen->SetTitle(L"�������� �����, ��� ��������� ����� ������ ��������");
        }



        



        if (iRequest <= 2 || iRequest == t_File_MailMerge_EEBase || iRequest == t_File_VacSchedule_EEBase)
        {
            COMDLG_FILTERSPEC fileTypes[] =
            {
                //{ L"All supported images", L"*.png;*.jpg;*.jpeg;*.psd" },
                { L"Excel files", L"*.xlsx" },

            };

            hr = pFileOpen->SetFileTypes(ARRAYSIZE(fileTypes), fileTypes);

        }
        else if (iRequest  == t_File_Backup_Folder || iRequest == t_File_MailMerge_AttachmentFolder || iRequest == t_File_Backup_FolderAlt)
        {
            pFileOpen->SetOptions(FOS_PICKFOLDERS | FOS_PATHMUSTEXIST);
        }
        else if (iRequest == t_File_Backup_Files || iRequest == t_File_MailMerge_AttachmentFiles)
        {
            pFileOpen->SetOptions(FOS_ALLOWMULTISELECT);
            iMultiSelect = 1;
        }

       //IShellItem* psi;
       //SHCreateItemFromParsingName(L"C:\\AMD", NULL, IID_IShellItem, (void**)&psi);
       //pFileOpen->SetFolder(psi);


        if (SUCCEEDED(hr))
        {
            // Show the Open dialog box.
            hr = pFileOpen->Show(NULL);

            // Get the file name from the dialog box.
            if (SUCCEEDED(hr))
            {

                if (iMultiSelect)
                {
                    IShellItemArray* pRets;
                    hr = pFileOpen->GetResults(&pRets);

                    if (SUCCEEDED(hr)) {

                        DWORD count;
                        pRets->GetCount(&count);
                        for (DWORD i = 0; i < count; i++) {
                            IShellItem* pRet;
                            LPWSTR nameBuffer;
                            pRets->GetItemAt(i, &pRet);
                            hr = pRet->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &nameBuffer);
                            if (SUCCEEDED(hr))
                            {
                                f_Open_File(iRequest, std::wstring(nameBuffer), 0);
                            }

                            pRet->Release();
                            CoTaskMemFree(nameBuffer);
                        }
                        pRets->Release();
                    }
                }else {

                    IShellItem* pItem;
                    hr = pFileOpen->GetResult(&pItem);
                    if (SUCCEEDED(hr))
                    {
                        PWSTR pszFilePath;
                        hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

                        // Display the file name to the user.
                        if (SUCCEEDED(hr))
                        {
                            f_Open_File(iRequest, pszFilePath, 0);
                            CoTaskMemFree(pszFilePath);
                        }

                    }
                     pFileOpen->Release();
                }
            }
            else {
                f_Clear_File(iRequest);
            }
            CoUninitialize();
        }
    }
    return 0;
}

void f_Save_File(int iRequest, wstring w_FilePath)
{
    if (iRequest == t_File_EElist_BaseToSave)
    {
        f_Excel_CreateExampleBase(w_FilePath);
    }

}




void f_Save_File_Dialog(int iRequest)
{

    HRESULT hr = CoInitializeEx(NULL, COINIT_APARTMENTTHREADED |
        COINIT_DISABLE_OLE1DDE);
    if (SUCCEEDED(hr))
    {
        IFileSaveDialog* pFileSave;
        hr = CoCreateInstance(CLSID_FileSaveDialog, NULL, CLSCTX_ALL,
            IID_IFileSaveDialog, reinterpret_cast<void**>(&pFileSave));

        if (iRequest == t_File_EElist_BaseToSave)
        {
            COMDLG_FILTERSPEC fileTypes[] =
            {
                { L"Excel files", L"*.xlsx" },

            };

            hr = pFileSave->SetFileTypes(ARRAYSIZE(fileTypes), fileTypes);

        }



        if (SUCCEEDED(hr))
        {

            // Show the Open dialog box.
            hr = pFileSave->Show(NULL);

            if (SUCCEEDED(hr))
            {
                IShellItem* pItem;
                hr = pFileSave->GetResult(&pItem);
                if (SUCCEEDED(hr))
                {
                    PWSTR pszFilePath;
                    hr = pItem->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);

                    // Display the file name to the user.
                    if (SUCCEEDED(hr))
                    {
                        f_Save_File(iRequest, pszFilePath);
                        CoTaskMemFree(pszFilePath);
                    }

                }

                pFileSave->Release();
            }

            CoUninitialize();
        }
    }
}