


void ExcelSerialDateToDMY(int nSerialDate, int& nDay, int& nMonth, int& nYear)
{
	// Modified Julian to DMY calculation with an addition of 2415019
	int l = nSerialDate + 68569 + 2415019;
	int n = int((4 * l) / 146097);
	l = l - int((146097 * n + 3) / 4);
	int i = int((4000 * (l + 1)) / 1461001);
	l = l - int((1461 * i) / 4) + 31;
	int j = int((80 * l) / 2447);
	nDay = l - int((2447 * j) / 80);
	l = int(j / 11);
	nMonth = j + 2 - (12 * l);
	nYear = 100 * (n - 49) + i + l;
}

int DMYToExcelSerialDate(int nDay, int nMonth, int nYear)
{
	// DMY to Modified Julian calculated with an extra subtraction of 2415019.
	return int((1461 * (nYear + 4800 + int((nMonth - 14) / 12))) / 4) +
		int((367 * (nMonth - 2 - 12 * ((nMonth - 14) / 12))) / 12) -
		int((3 * (int((nYear + 4900 + int((nMonth - 14) / 12)) / 100))) / 4) +
		nDay - 2415019 - 32075;
}







#include "pathcch.h"




bool f_Create_Folders_For_Path(char* s_FilePath)
{


	auto FolderPath = std::filesystem::path(s_FilePath).parent_path();

	if (FolderPath.string().length() > 0)
	{
		std::filesystem::create_directories(FolderPath);
	}
	return true;
}



std::wstring f_Excel_Get_WStringFromCell(XLCell cell)
{


	int iCellType = -1; // 0 for Number, 1 for string 

	switch (cell.value().type()) {

	case XLValueType::Integer:
		iCellType = 0;
		break;
	case XLValueType::String:
		iCellType = 1;
		break;

	default:
		break;
	}

	wstring s_empty = L" ";

	if (iCellType == 1)
	{

		string s_TempText_Get;
		s_TempText_Get = cell.value().get<string>();

		if (s_TempText_Get.length() > 0)
		{
			return f_string_to_wString(s_TempText_Get);
		}
		else {
			return s_empty;
		}
	}
	else if (iCellType == 0)
	{

		return to_wstring(cell.value().get<int64_t>());
	}
	else {

		return L"";
	}



}

std::wstring f_Excel_Get_Date_Wstring_FromCell(XLCell cell)
{
	int iCellType = -1; // 0 for Number, 1 for string 

	switch (cell.value().type()) {

	case XLValueType::Integer:
		iCellType = 0;
		break;
	case XLValueType::String:
		iCellType = 1;
		break;

	default:
		break;
	}

	wstring w_result{};

	if (iCellType == 0)
	{
		int iDay = 0, iMonth = 0, iYear = 0;
		int iDate = 0;

		iDate = cell.value().get<int64_t>();
		ExcelSerialDateToDMY(iDate, iDay, iMonth, iYear);

		if (iDay < 10)
		{
			w_result += to_wstring(0);
		}

		w_result = to_wstring(iDay) + L".";

		if (iMonth < 10)
		{
			w_result += to_wstring(0);
		}

		w_result += to_wstring(iMonth) + L"." + to_wstring(iYear);

	}
	else if (iCellType == 1)
	{
		string s_TempText_Get;
		s_TempText_Get = cell.value().get<string>();

		if (s_TempText_Get.length() > 0)
		{
			w_result = f_string_to_wString(s_TempText_Get);
		}
		else {
			w_result = L"Cannot extract date - String length is 0";
		}
	}
	else {
		w_result = L"Cannot extract date";
	}





	return w_result;


}






bool f_Excel_Process_File(Worker* ChosenWorker, int iRequest)
{
	

	g_FileUniqueCount += 1;
	wstring w_TempFilePath{};
	string s_TempFilePath{}; // for Excel library

	//std::vector<wstring> vec1;
	//boost::split_regex(vec1, w_TempFilePath, boost::regex(" "));

	if (iRequest == t_File_Reestr_Docs)
	{
		Frame_Main->m_text->AppendText(("\nЧтение реестра присланных документов"));
		w_TempFilePath = g_ReestrDocsPath;
		s_TempFilePath = f_wString_to_string(w_TempFilePath);
	}
	else if (iRequest == t_File_Reestr_Orders)
	{
		Frame_Main->m_text->AppendText(("\nЧтение реестра Приказов"));
		w_TempFilePath = g_ReestrOrdersPath;
		s_TempFilePath = f_wString_to_string(w_TempFilePath);
	}
	else if (iRequest == t_File_Reestr_Agreements)
	{
		Frame_Main->m_text->AppendText(("\nЧтение реестра доп. соглашений"));
		w_TempFilePath = g_ReestrAgreementsPath;
		s_TempFilePath = f_wString_to_string(w_TempFilePath);
	}
	else
	{
		Frame_Main->m_text->AppendText(("\niRequest не валидный."));
	}






	XLDocument doc{};
	XLWorksheet sheet{};
	XLCell cell{};

	int iTotalSheets = 0, iTotalRows = 0,iTotalColumns = 0;
	int iColOrderNumber = 0, iColOrderDate = 0, iColEmployeeName = 0, iColComment = 0, iColSigned = 0, iColInFolder = 0;
	int iColVacDateStart = 0, iColVacDateEnd = 0, iColDate = 0, iColSentToPayroll = 0;

	int iRow = 1, iCol = 1;
	int iCurrentSheetType = -1;

	wstring w_TempText{};
	wstring w_SheetName{};
	wstring w_EEName{};
	wstring w_NameRequest{};


	if (f_Does_file_exist(w_TempFilePath) == 0)
	{
		Frame_Main->m_text->AppendText(("\nФайла не существует."));
		f_SetBusy(0);
		return false;
	}


	doc.open(s_TempFilePath);

	iTotalSheets = doc.workbook().worksheetCount();
	w_NameRequest = ChosenWorker->EENameFull;

	ReestrSearch_wOutputDebug += L"\nПоиск по::" + w_NameRequest + L"\n";

	for (const auto& s_WorksSheetName : doc.workbook().worksheetNames())
	{

		w_SheetName = f_string_to_wString(s_WorksSheetName);
		ReestrSearch_wOutputDebug += L"Лист::" + w_SheetName + L"\n";


		iColOrderNumber = 0;
		iColOrderDate = 0;
		iColEmployeeName = 0;
		iColComment = 0;
		iColSigned = 0;
		iColInFolder = 0;
		iColSentToPayroll = 0;
		iColVacDateStart = 0;
		iColVacDateEnd = 0;
		iColDate = 0;


		sheet = doc.workbook().worksheet(s_WorksSheetName);


		iTotalRows = sheet.rowCount();
		iTotalColumns = sheet.columnCount();

		ReestrSearch_wOutputDebug += L"Столбцов::" + to_wstring(iTotalColumns) + L"\n";

		iCol = 1;

		while (iCol <= iTotalColumns)
		{
			ReestrSearch_wOutputDebug += L"Проверяю столбец::" + to_wstring(iCol) + L"\n";
			cell = sheet.cell(1, iCol);


			w_TempText = f_Excel_Get_WStringFromCell(cell);



			if (boost::iequals(w_TempText, L"Номер приказа"))
			{
				iColOrderNumber = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Дата приказа"))
			{
				iColOrderDate = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"ФИО"))
			{
				iColEmployeeName = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Комментарий"))
			{
				iColComment = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Подписан"))
			{
				iColSigned = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"В папке"))
			{
				iColInFolder = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Отправил в бухгалтерию"))
			{
				iColSentToPayroll = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Период отпуска С"))
			{
				iColVacDateStart = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Период отпуска ПО"))
			{
				iColVacDateEnd = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}

			else if (boost::iequals(w_TempText, L"Дата"))
			{
				iColDate = iCol;
				ReestrSearch_wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else {
				//w_TextOutPut += L"---" + (w_TempText)+L"::не найден" + L"---\n";
			}


			iCol += 1;
		}
		Frame_Main->m_text->AppendText("-1");


		iCurrentSheetType = -1;

		if (iRequest == t_File_Reestr_Orders)
		{

			if ((boost::icontains(w_SheetName, L"По личному составу")) || (boost::icontains(w_SheetName, L"Отпуска")))
			{
				if ((iColOrderNumber > 0) && (iColOrderDate > 0) && (iColEmployeeName > 0) && (iColComment > 0) && (iColSigned > 0) && (iColInFolder > 0) && (iColSentToPayroll > 0))
				{
					Frame_Main->m_text->AppendText("0");
					ReestrSearch_wOutputDebug += L"--НАШЁЛ СТОЛБЦЫ--::" + w_SheetName + "\n";
					iCurrentSheetType = t_Sheet_Reestr_Orders;
				}
				else {
					ReestrSearch_wOutputDebug += L"-- НЕ НАШЁЛ СТОЛБЦОВ--\n";
					continue;
				}
			}
			else if (boost::icontains(w_SheetName, L"Сверхурочка"))
			{
				if ((iColOrderNumber > 0) && (iColOrderDate > 0) && (iColEmployeeName > 0) && (iColSigned > 0))
				{
					ReestrSearch_wOutputDebug += L"--НАШЁЛ СТОЛБЦЫ--::" + w_SheetName + "\n";
					iCurrentSheetType = t_Sheet_Reestr_Overtime;
				}
				else {
					ReestrSearch_wOutputDebug += L"-- НЕ НАШЁЛ СТОЛБЦОВ--\n";
					continue;
				}

			}
			else {
				ReestrSearch_wOutputDebug += L"--ЛИСТ НЕ ПОДХОДИТ--\n";
				continue;
			}

		}
		else if ((iRequest == t_File_Reestr_Agreements))
		{
			if ((iColDate > 0) && (iColEmployeeName > 0) && (iColComment > 0) && (iColInFolder > 0))
			{
				ReestrSearch_wOutputDebug += L"--НАШЁЛ СТОЛБЦЫ--::" + w_SheetName + "\n";
				iCurrentSheetType = t_Sheet_Reestr_Overtime;
			}
			else {
				ReestrSearch_wOutputDebug += L"-- НЕ НАШЁЛ СТОЛБЦОВ--\n";
				continue;
			}

		}else
		{
			continue;
		}

		iRow = 2;

			while (iRow <= iTotalRows)
			{
				w_EEName = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColEmployeeName));
				Frame_Main->m_text->AppendText(w_SheetName.c_str());
				Frame_Main->m_text->AppendText(to_string(iRow));


					if (iRequest == t_File_Reestr_Orders)
					{
							if (iCurrentSheetType == t_Sheet_Reestr_Orders)
							{
								if (boost::icontains(w_EEName, w_NameRequest))
								{
									ReestrSearch_iFoundName = 1;
									w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColInFolder));
									if (!boost::iequals(w_TempText, L"Да"))
									{
										ReestrSearch_iFoundDocuments = 1;
										ReestrSearch_wOutputDebug += L"\n" + w_EEName + L"----НЕ В ПАПКЕ----\n";
										ReestrSearch_wOutputDebug += L"-- Приказ №" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColOrderNumber)) + L" от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColOrderDate)) + "\n";
										ReestrSearch_wOutputResult += (L"-- Приказ №" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColOrderNumber)) + L" от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColOrderDate)));
										w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColSigned));
										if (!boost::iequals(w_TempText, L"Да"))
											ReestrSearch_wOutputResult += L"(скана нет)";

										ReestrSearch_wOutputResult += L"\n";

									}else {

										ReestrSearch_wOutputDebug += L"--В папке--\n";
									}
								}
							}else if (iCurrentSheetType == t_Sheet_Reestr_Overtime)
							{
								if (boost::icontains(w_EEName, ChosenWorker->EENameLast))
								{
									ReestrSearch_iFoundName = 1;
									w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColSigned));
									if (!boost::iequals(w_TempText, L"Да"))
									{
										ReestrSearch_iFoundDocuments = 1;
										ReestrSearch_wOutputDebug += L"\n" + w_EEName + L"----НЕ В ПАПКЕ----\n";
										ReestrSearch_wOutputDebug += L"-- Приказ №" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColOrderNumber)) + L" от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColOrderDate)) + "\n";
										ReestrSearch_wOutputResult += (L"-- Приказ №" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColOrderNumber)) + L" от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColOrderDate)) + L"\n");
									}else {

										ReestrSearch_wOutputDebug += L"--В папке--\n";
									}
								}
							}

					}else if (iRequest == t_File_Reestr_Agreements)
					{
						if (boost::icontains(w_EEName, w_NameRequest))
						{
							ReestrSearch_iFoundName = 1;
							w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColInFolder));
							if (!boost::iequals(w_TempText, L"Да"))
							{
								ReestrSearch_iFoundDocuments = 1;
								ReestrSearch_wOutputDebug += L"\n" + w_EEName + L"----НЕ В ПАПКЕ----\n";
								ReestrSearch_wOutputDebug += L"-- Доп соглашение от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColDate)) + L", комментарий::" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColComment)) + "\n";
								ReestrSearch_wOutputResult += (L"-- Доп соглашение от " + f_Excel_Get_Date_Wstring_FromCell(sheet.cell(iRow, iColDate)) + L", комментарий::" + f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColComment)) + L"\n");
							}else {

								ReestrSearch_wOutputDebug += L"--В папке--\n";
							}
						}
					}

				iRow += 1;
			}
	}

	//if (iFoundName)
	//{
	//	if (iFoundDocuments)
	//	{

	//		Frame_Main->m_text->AppendText("Сотрудник найден.Приказы найдены.");
	//		if (g_DebugMode)
	//		{
	//			Frame_Main->m_text->AppendText(w_TextOutPut);
	//		}



	//		Frame_Main->m_text->AppendText("\nРезультат поиска::\n");
	//		Frame_Main->m_text->AppendText(w_TextOutPutResult);

	//	}
	//	else {
	//		Frame_Main->m_text->AppendText("Сотрудник найден.Приказов не найдено.");
	//		if (g_DebugMode)
	//		{
	//			Frame_Main->m_text->AppendText(w_TextOutPut);
	//		}
	//	}

	//}
	//else {
	//	Frame_Main->m_text->AppendText("По запросу работников не найдено.");
	//	if (g_DebugMode)
	//	{
	//		Frame_Main->m_text->AppendText(w_TextOutPut);
	//	}
	//}


	Frame_Main->m_text->AppendText(("\nProcessed::") + g_ReestrOrdersPath);
	doc.close();
	return true;
}


bool f_Excel_OpenEEBase()
{
	f_SetBusy(1);
	wstring w_TempFilePath{};
	string s_TempFilePath{}; // for Excel library

	w_TempFilePath = g_MailMergeEEBasePath;
	//w_TempFilePath = L"C:\\Users\\Alex\\Desktop\\База_Работников.xlsx";


	s_TempFilePath = f_wString_to_string(w_TempFilePath);

	XLDocument doc{};
	XLWorksheet sheet{};
	//XLCell cell{};


	int iFoundSheet = 0;

	int iTotalSheets = 0, iTotalRows = 0, iTotalColumns = 0;

	int iColEENameFull = 0;
	int iColEEEmailWork = 0;
	int iColEEEmailHome = 0;

	int iRow = 1, iCol = 1;

	wstring w_TempText{};
	wstring w_SheetName{};
	wstring w_EEName{};
	wstring wOutputDebug{};

	Worker* NewWorker = NULL;

	vector<int> CustomColumnsArray{};
	wstring wCustomColumnHeader{}, wCustomColumnValue{};


	if (f_Does_file_exist(w_TempFilePath) == 0)
	{
		Frame_Main->m_text->AppendText(("\nФайла не существует."));
		return false;
	}

	Frame_Main->m_text->AppendText(("\nНачинаю импорт базы"));

	g_WorkerArray.clear();

	doc.open(s_TempFilePath);
	iTotalSheets = doc.workbook().worksheetCount();

	for (const auto& s_WorksSheetName : doc.workbook().worksheetNames())
	{

		iColEENameFull = 0;
		iColEEEmailWork = 0;
		iColEEEmailHome = 0;
		w_SheetName = f_string_to_wString(s_WorksSheetName);
		sheet = doc.workbook().worksheet(s_WorksSheetName);
		wOutputDebug += L"Лист::" + w_SheetName + L"\n";
		iTotalRows = sheet.rowCount();
		iTotalColumns = sheet.columnCount();


		iCol = 1;
		while (iCol <= iTotalColumns)
		{
			wOutputDebug += L"Проверяю столбец::" + to_wstring(iCol) + L"\n";
			w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(1, iCol));


			if (boost::iequals(w_TempText, L"ФИО"))
			{
				iColEENameFull = iCol;
				wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}else if (boost::iequals(w_TempText, L"Email_Work"))
			{
				iColEEEmailWork = iCol;
				wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}
			else if (boost::iequals(w_TempText, L"Email_Home"))
			{
				iColEEEmailHome = iCol;
				wOutputDebug += L"---" + (w_TempText)+L"::" + to_string(iCol) + L"---\n";
			}else{

				if (w_TempText.length() > 0) // if first row has any text
				{
					CustomColumnsArray.push_back(iCol);
				}

			}
			iCol++;
		}


		if (iColEENameFull > 0  && iColEEEmailWork > 0 && iColEEEmailHome > 0)
		{
			wOutputDebug += L"--НАШЁЛ СТОЛБЦЫ--::" + w_SheetName + "\n";
			iFoundSheet = 1;
			break;

		}
		else {
			wOutputDebug += L"-- НЕ НАШЁЛ СТОЛБЦЫ--::" + w_SheetName + "\n";
		}



	}


	if (!iFoundSheet)
	{
		Frame_Main->m_text->AppendText(("\nВ базе нет подходящих листов"));
		return false;
	}

	iRow = 2;


	while (iRow <= iTotalRows)
	{
		w_EEName = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColEENameFull));

		if (w_EEName.length() > 0)
		{
			NewWorker = new Worker;
			NewWorker->EENameFull = w_EEName;


			std::vector<wstring> vec1;
			boost::split_regex(vec1, w_EEName, boost::regex(" "));
			NewWorker->EENameLast = vec1[0];
		

			w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColEEEmailWork));

			if (w_TempText.length() > 0)
			{
				NewWorker->EEEmailWork = w_TempText;
			}
			else {
				NewWorker->EEEmailWork = L"";
			}

			w_TempText = f_Excel_Get_WStringFromCell(sheet.cell(iRow, iColEEEmailHome));

			if (w_TempText.length() > 0)
			{
				NewWorker->EEEmailHome = w_TempText;
			}
			else {
				NewWorker->EEEmailHome = L"";
			}

			for (vector<int>::iterator it = CustomColumnsArray.begin(); it != CustomColumnsArray.end(); ++it)
			{
				wCustomColumnHeader = f_Excel_Get_WStringFromCell(sheet.cell(1, Iter));
				wCustomColumnHeader = L"<" + std::move(wCustomColumnHeader) + L">";
				wCustomColumnValue = f_Excel_Get_WStringFromCell(sheet.cell(iRow, Iter));

				//NewWorker->MergeFieldsMap.insert(f_Excel_Get_WStringFromCell(sheet.cell(1, Iter)), f_Excel_Get_WStringFromCell(sheet.cell(iRow, Iter)));
				NewWorker->MergeFieldsMap.insert(make_pair(wCustomColumnHeader, wCustomColumnValue));
			}


			//NewWorker->MergeFieldsMap.insert(w_EEName, w_EEName);







			g_WorkerArray.push_back(NewWorker);
		}

		iRow++;
	}


		wxArrayString ar_list;
		for (std::vector<Worker*>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
		{
			Frame_Main->m_text->AppendText(Iter->EENameFull);
			Frame_Main->m_text->AppendText(L"/n");

			ar_list.Add(Iter->EENameFull);
		}

		Frame_Main->m_EEList->InsertItems(ar_list, 0);


		


	Frame_Main->m_text->AppendText(("\nИмпорт базы завершён."));
	f_SetBusy(0);
	return true;
}



void f_Excel_CreateExampleBase(wstring wFileName)
{

	wFileName = wFileName + ".xlsx";

	string s_TempFilePath{}; // for Excel library
	s_TempFilePath = f_wString_to_string(wFileName);


	XLDocument doc;
	doc.create(s_TempFilePath);
	auto wks = doc.workbook().worksheet("Sheet1");

	wks.cell(1,1).value() = f_wString_to_string(L"ФИО");
	wks.cell(1,2).value() = "Email_Work";
	wks.cell(1,3).value() = "Email_Home";

	doc.save();
}


