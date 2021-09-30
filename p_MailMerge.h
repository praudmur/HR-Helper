void f_MailMerge_Sent_Letter(wstring Subject, wstring EmailWork, wstring EmailHome, wstring Body, vector<wstring> w_Attachments)
{

    Outlook::_ApplicationPtr m_OutlookAppPtr;
    HRESULT hr;
    hr = m_OutlookAppPtr.CreateInstance(__uuidof(Outlook::Application));



    IDispatchPtr IMailItemDispatchPtr;
    IMailItemDispatchPtr = m_OutlookAppPtr->CreateItem(Outlook::olMailItem);


    Outlook::_MailItemPtr IMailItemPtr;
    Outlook::_MailItem* IMailItem = NULL;
    hr = IMailItemDispatchPtr.QueryInterface(__uuidof(Outlook::_MailItem), &IMailItem);

    IMailItemPtr.Attach(IMailItem);




    IMailItemPtr->Subject = Subject.c_str();
    IMailItemPtr->To = EmailWork.c_str();
    IMailItemPtr->Body = Body.c_str();
    IMailItemPtr->CC = EmailHome.c_str();


        for (std::vector<wstring>::iterator it = w_Attachments.begin(); it != w_Attachments.end(); ++it)
        {
            if (f_Does_file_exist(Iter))
                IMailItemPtr->Attachments->Add(Iter.c_str());
        }



    IMailItemPtr->Save();
    if (g_MailMerge_ShowLetters)
    {
        IMailItemPtr->Display();
    }


}





void MyFrame::OnMailMergeButtonPressed(wxCommandEvent& WXUNUSED(event))
{

    f_SetBusy(1);

    int iValue = -1, iPos = 0, iCount = -1;



    m_text->AppendText("Mail Merge Button pressed\n");

    wxArrayInt CheckedItems;
    m_EEList->GetCheckedItems(CheckedItems);

    iCount = CheckedItems.GetCount();


    if (iCount <= 0)
    {
        m_text->AppendText("No EEspecified\n");
        f_SetBusy(0);
        return;
    }

    wstring temp_text{};
    wstring w_LetterBody{};
    wstring w_LetterSubject{};
    vector<wstring> wAttachementsChecked{};
    vector<wstring> wAttachements{};
    vector<wstring> wToRemove{};
    wstring w_attachment_temp{};

    wxArrayString items;
    wstring w_tempString{};

    int iAttacmentFolderExists = f_Does_file_exist(g_MailMergeAttachmentFolderPath);


    if (g_MailMerge_AddStaticAttachments)
    {
        Frame_Main->m_MailMerge_ListBoxAttachments->GetStrings(items);
        int num = items.GetCount();
        if ((num > 0))
        {

            for (int i = 0; i < num; i++)
            {
                w_tempString = items[i];

                if (w_tempString.length() > 0 && f_Does_file_exist(w_tempString))
                {
                    wAttachementsChecked.push_back(w_tempString);
                }
            }
        }
    }




    while (iPos < iCount)
    {
        iValue = CheckedItems[iPos];

        for (std::vector<Worker*>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
        {
            m_text->AppendText(L"\nComparing::" + Frame_Main->m_EEList->GetString(iValue) + L" with " + Iter->EENameFull);

            temp_text = Frame_Main->m_EEList->GetString(iValue);
            if (boost::iequals(temp_text, Iter->EENameFull))
            {
                m_text->AppendText(L"\nSENDING");
                w_LetterBody = m_MailMerge_Text_LetterBody->GetValue();
                w_LetterSubject = m_MailMerge_Text_LetterSubject->GetValue();

                wAttachements.clear();
                wToRemove.clear();
                wAttachements.insert(std::end(wAttachements), std::begin(wAttachementsChecked), std::end(wAttachementsChecked));


                if (g_MailMerge_LookforAttachments)
                {
                    if (iAttacmentFolderExists)
                    {
                        for (auto const& dir_entry : std::filesystem::directory_iterator{ g_MailMergeAttachmentFolderPath })
                        {
                            if (boost::icontains(dir_entry.path().filename().c_str(), Iter->EENameFull))
                                wAttachements.push_back(dir_entry.path());
                        }
                    }
                }

                f_MailMerge_ReplaceFields(&w_LetterBody, &w_LetterSubject, Iter);

                if (iAttacmentFolderExists)
                {

                    boost::wregex expression(L"<вложить>([^<]*)</вложить>");

                    boost::regex_token_iterator<std::wstring::iterator> w_it{ w_LetterBody.begin(), w_LetterBody.end(),expression, 1 };
                    boost::regex_token_iterator<std::wstring::iterator> end;
                    while (w_it != end)
                    {
                        w_attachment_temp = (*w_it).str();
                        for (auto const& dir_entry : std::filesystem::directory_iterator{ g_MailMergeAttachmentFolderPath })
                        {
                            if (boost::icontains(dir_entry.path().filename().c_str(), w_attachment_temp))
                            {
                                wAttachements.push_back(dir_entry.path());
                            }
                        }

                        w_attachment_temp = L"<вложить>" + std::move(w_attachment_temp) + L"</вложить>";
                        wToRemove.push_back(w_attachment_temp);
                        w_it++;
                    }

                    for (std::vector<wstring>::iterator it_W_Remove = wToRemove.begin(); it_W_Remove != wToRemove.end(); ++it_W_Remove)
                    {
                        boost::ireplace_all(w_LetterBody, (*it_W_Remove), L"");
                    }

                }

                f_MailMerge_Sent_Letter(w_LetterSubject, Iter->EEEmailWork, Iter->EEEmailHome, w_LetterBody, wAttachements);
                break;
            }

        }


        iPos++;
    }

    f_SetBusy(0);
}

void MyFrame::OnMailMergeAddAttachmentFolder(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening Mail Merge Attachment Folder \n");
    f_Open_File_Dialog(t_File_MailMerge_AttachmentFolder);
}

void MyFrame::OnMailMergeAddAttachmentFiles(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening Mail Merge Attachment Files \n");
    f_Open_File_Dialog(t_File_MailMerge_AttachmentFiles);
}