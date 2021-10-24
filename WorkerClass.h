



class Worker
{
public:
	wstring EENameFull{};
	wstring EEDepartment{};
	wstring EENameLast{};
	wstring EEEmailWork{};
	wstring EEEmailHome{};
	map <wstring, wstring> MergeFieldsMap;

};

vector<Worker*> g_WorkerArray;




/// ////////////////////VACATIONS




class Dates
{
public:
	int iDateFrom = -1;
	int iDateTo = -1;
};



class Vacation
{
public:
	int iWorkerID = -1;
	int iOrderDate = -1;
	vector<Dates*> VacDates;
	int iOnSchedule = -1;
	vector<Dates*> VacDatesAdded;
	int iTransferred = -1;
	vector<Dates*> VacDatesTransferFrom;
};



class WorkerVac
{
public:
	int iID = -1;
	wstring EENameFull{}; 
};

vector<WorkerVac*> g_WorkerVacArray;
vector<Vacation*> g_VacationsArray;


wstring f_WorkerVacNameFromID(int iID)
{
	for (std::vector<WorkerVac*>::iterator it = g_WorkerVacArray.begin(); it != g_WorkerVacArray.end(); ++it)
	{
		if (Iter->iID == iID)
		{
			return  (Iter->EENameFull);
		}
	}
	
	return (L"EE not found");
}




//class WorkerEE
//{
//public:
//	wstring EENameFull{};
//	wstring EENameLast{};
//	wstring EEEmailWork{};
//	wstring EEEmailHome{};
//	map <wstring, wstring> MergeFieldsMap;
//
//};
//
//vector<WorkerEE*> g_WorkerArray;



//int f_Worker_GetKey(int iNumber = 0, int iRequest = 0)
//{
//	//Request
//	//0 - tab nomer
//	//1 - SSO
//
//	int iFound = 0;
//	int iIndex;
//
//	if (iRequest)
//	{
//		for (std::vector<Worker>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
//		{
//			if ((*it).iSSO == iNumber)
//			{
//				iFound = 1;
//				iIndex = it - g_WorkerArray.begin();
//				break;
//			}
//
//		}
//
//	}
//	else {
//		for (std::vector<Worker>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
//		{
//			if ((*it).iTabelNomer == iNumber)
//			{
//				iFound = 1;
//				iIndex = it - g_WorkerArray.begin();
//				break;
//			}
//
//		}
//	}
//
//	if (iFound == 1)
//	{
//		return iIndex;
//	}
//	else { return -1; }
//
//
//
//}