

void f_VacSchedule_Run_Selection_For(wstring w_Chosen_WorkerName)
{
	Frame_Main->m_VacSchedule_EEVacList->ClearAll();


	Frame_Main->m_VacSchedule_EEVacList->InsertColumn(0, "Отпуск С", wxLIST_FORMAT_LEFT);
	Frame_Main->m_VacSchedule_EEVacList->InsertColumn(1, "Отпуск ДО", wxLIST_FORMAT_LEFT);
	wxListItem* item = new wxListItem();

	Frame_Main->m_VacSchedule_EEVacList->SetColumnWidth(0, 150);
	Frame_Main->m_VacSchedule_EEVacList->SetColumnWidth(1, 150);


	int iIndex = 0;

	for (std::vector<WorkerVac*>::iterator it = g_WorkerVacArray.begin(); it != g_WorkerVacArray.end(); ++it)
	{
		if (Iter->EENameFull == w_Chosen_WorkerName)
		{
			for (std::vector<Vacation*>::iterator it2 = g_VacationsArray.begin(); it2 != g_VacationsArray.end(); ++it2)
			{
				if (Iter2->iWorkerID == Iter->iID)
				{
					Frame_Main->m_VacSchedule_EEVacList->InsertItem(iIndex, *item);

					for (std::vector<Dates*>::iterator it3 = Iter2->VacDates.begin(); it3 != Iter2->VacDates.end(); ++it3)
					{
						Frame_Main->m_VacSchedule_EEVacList->SetItem(iIndex, 0, f_ExcelDateToWString(Iter3->iDateFrom));
						Frame_Main->m_VacSchedule_EEVacList->SetItem(iIndex, 1, f_ExcelDateToWString(Iter3->iDateTo));
					}

					iIndex++;
				}
			}

		}
	}
}