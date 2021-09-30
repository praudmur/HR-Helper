





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

        Frame_Main->m_text->AppendText("\nФайл загружен::");
        Frame_Main->m_text->AppendText(g_ReestrDocsPath);
        s_Message = "Реестр приказов>>>>>" + g_ReestrDocsPath;

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

        Frame_Main->m_text->AppendText("\nФайл загружен::");
        Frame_Main->m_text->AppendText(g_ReestrOrdersPath);
        s_Message = "Реестр приказов>>>>>" + g_ReestrOrdersPath;

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

        Frame_Main->m_text->AppendText("\nФайл загружен::");
        Frame_Main->m_text->AppendText(g_ReestrAgreementsPath);
        s_Message = "Реестр приказов>>>>>" + g_ReestrAgreementsPath;

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

    }else if (iRequest == t_File_Backup_Folder)
    {

        g_BackupFolderPath = w_FilePath;
        s_Message = "Папка резервной копии>>>" + g_BackupFolderPath;

        Frame_Main->m_Backup_Status_BackupFolderPath->SetLabelText(s_Message);
        Frame_Main->m_Backup_Status_BackupFolderPath->SetBackgroundColour(*wxGREEN);

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadBackupFolderPath", w_FilePath);
        }

    }else if (iRequest == t_File_MailMerge_EEBase)
    {

        g_MailMergeEEBasePath = w_FilePath;

        s_Message = "База работников Mail Merge>>>" + g_MailMergeEEBasePath;
        Frame_Main->m_text->AppendText(s_Message);

        f_Excel_OpenEEBase();

        if (!bFromINI)
        {
            f_SetINIValue("Settings", "bLoadMailMergeBaseEEFile", w_FilePath);
        }

    }else if (iRequest == t_File_MailMerge_AttachmentFolder)
    {

        g_MailMergeAttachmentFolderPath = w_FilePath;
        s_Message = "Папка вложений>>>" + g_MailMergeAttachmentFolderPath;

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
            pFileOpen->SetTitle(L"Выберите реестр получения документов");
        }
        else if (iRequest == t_File_Reestr_Orders)
        {
            pFileOpen->SetTitle(L"Выберите реестр приказов");
        }else if (iRequest == t_File_Reestr_Agreements)
        {
            pFileOpen->SetTitle(L"Выберите реестр доп. соглашений");
        }
        else if (iRequest == t_File_Backup_Files)
        {
            pFileOpen->SetTitle(L"Выберите файлы, которые нужно скопировать");
        }else if (iRequest == t_File_Backup_Folder)
        {
            pFileOpen->SetTitle(L"Выберите папку, куда нужно скопировать файлы");
        }
        else if (iRequest == t_File_MailMerge_EEBase)
        {
            pFileOpen->SetTitle(L"Выберите файл базы работников");
        }else if (iRequest == t_File_MailMerge_AttachmentFolder)
        {
            pFileOpen->SetTitle(L"Выберите папку, где программа будет искать вложения");
        }



        



        if (iRequest <= 2 || iRequest == t_File_MailMerge_EEBase)
        {
            COMDLG_FILTERSPEC fileTypes[] =
            {
                //{ L"All supported images", L"*.png;*.jpg;*.jpeg;*.psd" },
                { L"Excel files", L"*.xlsx" },

            };

            hr = pFileOpen->SetFileTypes(ARRAYSIZE(fileTypes), fileTypes);

        }
        else if (iRequest  == t_File_Backup_Folder || iRequest == t_File_MailMerge_AttachmentFolder)
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