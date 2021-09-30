



class Worker
{
public:
	wstring EENameFull{};
	wstring EENameLast{};
	wstring EEEmailWork{};
	wstring EEEmailHome{};
	map <wstring, wstring> MergeFieldsMap;

};

vector<Worker*> g_WorkerArray;




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