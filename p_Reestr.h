


void f_Reestr_Search_Start()
{
    f_SetBusy(1);

    if (g_FilesLoaded == 0)
    {
        Frame_Main->m_text->AppendText("\n��� ����������� ������!");
        f_SetBusy(0);
        return;
    }

    Worker* ChosenWorker = NULL;
    wstring w_request{};

    int iValue = -1, iPos = 0, iCount = -1;

    wxArrayInt CheckedItems;
    Frame_Main->m_EEList->GetCheckedItems(CheckedItems);
    iCount = CheckedItems.GetCount();


    if (iCount <= 0)
    {
        Frame_Main->m_text->AppendText("\n��������� �� �������.");
        f_SetBusy(0);
        return;
    }

    while (iPos < iCount)
    {
        iValue = CheckedItems[iPos];

        for (std::vector<Worker*>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
        {

            if (boost::iequals((wstring)Frame_Main->m_EEList->GetString(iValue), Iter->EENameFull))
            {
                ChosenWorker = Iter;
                break;
            }
        }

        iPos++;
    }


    if (ChosenWorker == NULL)
    {
        Frame_Main->m_text->AppendText("\n������ ������.");
        f_SetBusy(0);
        return;
    }
    else {
        Frame_Main->m_text->AppendText("\n����� ���������");
    }

    w_request = ChosenWorker->EENameFull;

    ReestrSearch_iFoundName = 0;
    ReestrSearch_iFoundDocuments = 0;
    ReestrSearch_wOutputResult.clear();
    ReestrSearch_wOutputDebug.clear();

    Frame_Main->m_text->AppendText("\n������� ����� �� " + w_request);


    if (g_ReestrOrdersPathLoaded)
        f_Excel_Process_File(ChosenWorker, t_File_Reestr_Orders);

    if (g_ReestrDocsPathLoaded)
        f_Excel_Process_File(ChosenWorker, t_File_Reestr_Docs);

    if (g_ReestrAgreementsPathLoaded)
        f_Excel_Process_File(ChosenWorker, t_File_Reestr_Agreements);





    if (ReestrSearch_iFoundName)
    {
        if (ReestrSearch_iFoundDocuments)
        {

            Frame_Main->m_text->AppendText("��������� ������.������� �������.");
            if (g_DebugMode)
            {
                Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
            }



            Frame_Main->m_text->AppendText("\n��������� ������::\n");
            Frame_Main->m_text->AppendText(ReestrSearch_wOutputResult);

        }
        else {
            Frame_Main->m_text->AppendText("��������� ������.�������� �� �������.");
            if (g_DebugMode)
            {
                Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
            }
        }

    }
    else {
        Frame_Main->m_text->AppendText("�� ������� ���������� �� �������.");
        if (g_DebugMode)
        {
            Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
        }
    }


    f_SetBusy(0);

}


void f_Reestr_Mass_Reminder_Start()
{
 
    Frame_Main->m_text->AppendText("\n1");
    f_SetBusy(1);
    Frame_Main->m_text->AppendText("\n2");
    int iFound = 0;
    if (g_FilesLoaded == 0)
    {
        Frame_Main->m_text->AppendText("\n��� ����������� ������!");
        f_SetBusy(0);
        return;
    }
    Frame_Main->m_text->AppendText("\n3");

    wstring w_request{}, temp_text{}, w_LetterBody{};
    vector<wstring> wAttachements{};

    int iValue = -1, iPos = 0, iCount = -1;

    wxArrayInt CheckedItems;
    Frame_Main->m_EEList->GetCheckedItems(CheckedItems);
    iCount = CheckedItems.GetCount();

    Worker* ChosenWorker = NULL;
    Frame_Main->m_text->AppendText("\n4");
    if (iCount <= 0)
    {
        Frame_Main->m_text->AppendText("\n��������� �� �������.");
        f_SetBusy(0);
        return;
    }
    Frame_Main->m_text->AppendText("\n5");
    while (iPos < iCount)
    {
        Frame_Main->m_text->AppendText("\n6");
        iValue = CheckedItems[iPos];
        for (std::vector<Worker*>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
        {
            Frame_Main->m_text->AppendText("\n7");
            if (boost::iequals((wstring)Frame_Main->m_EEList->GetString(iValue), Iter->EENameFull))
            {
                Frame_Main->m_text->AppendText("\n8");
                ReestrSearch_iFoundName = 0;
                ReestrSearch_iFoundDocuments = 0;
                ReestrSearch_wOutputResult.clear();
                ReestrSearch_wOutputDebug.clear();
                ChosenWorker = Iter;

                Frame_Main->m_text->AppendText("\n������� ����� �� " + ChosenWorker->EENameFull);


                if (g_ReestrOrdersPathLoaded)
                    f_Excel_Process_File(ChosenWorker, t_File_Reestr_Orders);

                if (g_ReestrDocsPathLoaded)
                    f_Excel_Process_File(ChosenWorker, t_File_Reestr_Docs);

                if (g_ReestrAgreementsPathLoaded)
                    f_Excel_Process_File(ChosenWorker, t_File_Reestr_Agreements);





                if (ReestrSearch_iFoundName)
                {
                    if (ReestrSearch_iFoundDocuments)
                    {

                        Frame_Main->m_text->AppendText("��������� ������.������� �������.");
                        if (g_DebugMode)
                        {
                            Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
                        }



                        Frame_Main->m_text->AppendText("\n��������� ������::\n");
                        Frame_Main->m_text->AppendText(ReestrSearch_wOutputResult);


                        w_LetterBody = L"<���>, ������������!\n";
                        w_LetterBody += L"� ������ �������� ������� ���������� � ����� ������ ����, ������� ������ ����������, ������� � ���� ��� � �����:\n";
                        w_LetterBody += ReestrSearch_wOutputResult + "\n";
                        w_LetterBody += L"����������, �������� ��� �� �� ��������� ��������� �������.\n";
                        w_LetterBody += L"���� �� ������ ����� �� � ����� - �������� ���, � �������.\n";
                        w_LetterBody += L"� ���������,.\n";
                        w_LetterBody += L"���������";
                        f_MailMerge_ReplaceFields(&w_LetterBody, NULL, Iter);

                        f_MailMerge_Sent_Letter(L"���������", Iter->EEEmailWork, Iter->EEEmailHome, w_LetterBody, wAttachements);


                    }
                    else {
                        Frame_Main->m_text->AppendText("��������� ������.�������� �� �������.");
                        if (g_DebugMode)
                        {
                            Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
                        }
                    }

                }
                else {
                    Frame_Main->m_text->AppendText("�� ������� ���������� �� �������.");
                    if (g_DebugMode)
                    {
                        Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
                    }
                }





            }
        }


        iPos++;
    }







    if (ReestrSearch_iFoundName)
    {
        if (ReestrSearch_iFoundDocuments)
        {

            Frame_Main->m_text->AppendText("��������� ������.������� �������.");
            if (g_DebugMode)
            {
                Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
            }



            Frame_Main->m_text->AppendText("\n��������� ������::\n");
            Frame_Main->m_text->AppendText(ReestrSearch_wOutputResult);

        }
        else {
            Frame_Main->m_text->AppendText("��������� ������.�������� �� �������.");
            if (g_DebugMode)
            {
                Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
            }
        }

    }
    else {
        Frame_Main->m_text->AppendText("�� ������� ���������� �� �������.");
        if (g_DebugMode)
        {
            Frame_Main->m_text->AppendText(ReestrSearch_wOutputDebug);
        }
    }


    f_SetBusy(0);

}




void MyFrame::OnReestrButtonMassReminder(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("MASS REMINDER STARTS");
    f_Reestr_Mass_Reminder_Start();
    //return;
}

void MyFrame::OnReestrButtonPressed(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("REESTR  STARTS");
    f_Reestr_Search_Start();
}

