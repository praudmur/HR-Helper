





// For compilers that support precompilation, includes "wx/wx.h".
#include "wx/wxprec.h"


// for all others, include the necessary headers (this file is usually all you
// need because it includes almost all "standard" wxWidgets headers)
#ifndef WX_PRECOMP
#include "wx/wx.h"
#endif 


#include "wx/uiaction.h"


#include "wx/stopwatch.h"

#include "wx/datectrl.h"
#include "wx/dateevt.h"

#include "wx/editlbox.h"
#include "wx/listctrl.h"

#include <wx/app.h>
// ----------------------------------------------------------------------------
// resources
// ----------------------------------------------------------------------------

#define wxUSE_OWNER_DRAWN

// the application icon (under Windows it is in resources and even
// though we could still include the XPM here it would be unused)
#ifndef wxHAS_IMAGES_IN_RESOURCES
#include "../sample.xpm"
#endif

// ----------------------------------------------------------------------------
// constants
// ----------------------------------------------------------------------------

// IDs for the controls and the menu commands
enum
{
    // menu items
    wx_reestr_OpenReestrOrders = 1,
    wx_reestr_OpenReestrDocs, // Base upload indicator
    wx_reestr_Search_EmployeeName, // Search for name
    wx_reestr_OpenReestrAgreements,
    wx_reestr_StartSearchingReestrs,
    wx_reestr_StartMassReminder,



    //Date Diff panel
    wx_DateDiff_ExecuteDateDifference,

    //Backup panel
    wx_Backup_ExecuteBackup,
    wx_Backup_AddFolder,
    wx_Backup_AddFolderAlt,
    wx_Backup_AddFiles,
    wx_Backup_SaveFilesToINI,
    wx_Backup_ListBox,

    //Work EXP panel
    //wx_WorkExp_ExecuteWorkExp,

    //Main Merge panel
    wx_MailMerge_ExecuteMailMerge,
    wx_MailMerge_TextLetterBody,
    wx_MailMerge_AddAttachmentFolder,
    wx_MailMerge_AddAttachmentFiles,
    wx_MailMerge_ListBox,



    //Vac Schedule EE List
    wx_VacSchedule_EEList,
    wx_VacSchedule_Reestr,
    wx_VacSchedule_Menu_OpenBase,
    wx_VacSchedule_EEVacList,



    //EE list
    wx_EElist_AddEEBase,
    wx_EElist_SelectAll,
    wx_EElist_SelectNone,
    wx_EElist_CreateExampleBase,
    wx_EElist_Department,


    wx_chkBox_All,

    wx_SwitchPanelsReestr,
    wx_SwitchPanelsDateDifference,
    wx_SwitchPanelsBackup,
    wx_SwitchPanelsMailMerge,
    wx_SwitchPanelsVacSchedule,


    wx_Global_OnKeyPress,
    wx_OnLicence,
    wx_OnProgramWebsite,
    wx_Global_LastButton,

};

// ----------------------------------------------------------------------------
// private classes
// ----------------------------------------------------------------------------

// Define a new application type, each program should derive a class from wxApp
class MyApp : public wxApp
{
public:
    virtual bool OnInit() wxOVERRIDE;
    virtual int OnExit() wxOVERRIDE;
};




// Define a new frame type: this is going to be our main frame
class MyFrame : public wxFrame
{
public:
    // ctor(s)
    MyFrame(const wxString& title);

    void OnReestrButtonPressed(wxCommandEvent& event);
    void OnOpenReestrDocs(wxCommandEvent& event);
    void OnOpenReestrOrders(wxCommandEvent& event);
    void OnOpenReestrAgreements(wxCommandEvent& event);
    void OnReestrButtonMassReminder(wxCommandEvent& event);



    void OnDateDiffButtonPressed(wxCommandEvent& event);

    void OnBackupButtonPressed(wxCommandEvent& event);
    void OnBackupButtonAddFolder(wxCommandEvent& event);
    void OnBackupButtonAddFolderAlt(wxCommandEvent& event);
    void OnBackupButtonAddFiles(wxCommandEvent& event);
    void OnBackupButtonSaveFilesToINI(wxCommandEvent& event);


    void OnMailMergeButtonPressed(wxCommandEvent& event);
    void OnMailMergeAddAttachmentFolder(wxCommandEvent& event);
    void OnMailMergeAddAttachmentFiles(wxCommandEvent& event);

    void OnEElistAddBase(wxCommandEvent& event);
    void OnEElistSelectAll(wxCommandEvent& event);
    void OnEElistSelectNone(wxCommandEvent& event);
    void OnEElistCreateExampleBase(wxCommandEvent& event);
    void OnEElistDepartmentSelected(wxCommandEvent& event);


    void OnSwitchPanelsButton(wxCommandEvent& event);



    void OnVacScheduleOpenEEBase(wxCommandEvent& event);


    void OnVacScheduleEEListSelected(wxCommandEvent& event);




    void OnKeyPress(wxKeyEvent& event);

    void OnCheckBox(wxCommandEvent& event);

    void OnExit(wxCommandEvent& WXUNUSED(event)) { Close(1); }
    void OnAbout(wxCommandEvent& event);
    void OnLicence(wxCommandEvent& event);
    void OnProgramWebsite(wxCommandEvent& event);

    void OnMouseDoubleClick(wxMouseEvent& event);

    //Reestr buttons
    wxTextCtrl* m_text = NULL;
    wxTextCtrl* m_reestr_Status_ReestrOrdersPath = NULL;
    wxTextCtrl* m_reestr_Status_ReestrDocsPath = NULL;
    wxTextCtrl* m_reestr_Status_ReestrAgreementsPath = NULL;
    //wxTextCtrl* m_reestr_Search_EmployeeName = NULL;
    wxButton* m_reestr_Button_Search = NULL;
    wxButton* m_reestr_Button_Mass_Reminder = NULL;

    //Date diff buttons
    wxDatePickerCtrl* m_dateDiff_DatePickerFrom = NULL;
    wxDatePickerCtrl* m_dateDiff_DatePickerTo = NULL;
    wxButton* m_dateDiff_Button_Execute = NULL;

    //Backup buttons
    wxButton* m_Backup_Button_Execute = NULL;
    wxButton* m_Backup_Button_AddFolder = NULL;
    wxButton* m_Backup_Button_AddFolderAlt = NULL;
    wxButton* m_Backup_Button_AddFiles = NULL;
    wxButton* m_Backup_Button_SaveFilesToINI = NULL;
    wxTextCtrl* m_Backup_Status_BackupFolderPath = NULL;
    wxTextCtrl* m_Backup_Status_BackupFolderPathAlt = NULL;
    wxCheckBox* m_Backup_Checkbox_OpenFolder = NULL;
    wxEditableListBox* m_Backup_ListBoxFiles = NULL;

    //WorkEXP buttons
    //wxButton* m_WorkExp_Button_Execute = NULL;

    //Mail merge buttons
    wxButton* m_MailMerge_Button_Execute = NULL;
    wxButton* m_MailMerge_Button_AddAttachmentFolder = NULL;
    wxTextCtrl* m_MailMerge_Status_AttachmentFolderPath = NULL;
    wxTextCtrl* m_MailMerge_Text_LetterSubject = NULL;
    wxTextCtrl* m_MailMerge_Text_LetterBody = NULL;
    wxCheckBox* m_MailMerge_Checkbox_LookforAttachments = NULL;
    wxCheckBox* m_MailMerge_Checkbox_SendOnlyOneLetter = NULL;
    wxCheckBox* m_MailMerge_Checkbox_ShowLetters = NULL;
    wxCheckBox* m_MailMerge_Checkbox_LookOnlyByAttachments = NULL;
    wxCheckBox* m_MailMerge_Checkbox_AddStaticAttachments = NULL;
    wxEditableListBox* m_MailMerge_ListBoxAttachments = NULL;
    wxButton* m_MailMerge_Button_AddAttachmentFiles = NULL;


    //Vac Schedule EE List
    wxListBox* m_VacSchedule_EEList = NULL;
    wxListCtrl* m_VacSchedule_Reestr = NULL;
    wxListCtrl* m_VacSchedule_EEVacList = NULL;



    //EE list buttons
    wxButton* m_EEList_Button_AddBase = NULL;
    wxCheckListBox* m_EEList = NULL;
    wxCheckListBox* m_EEListDepartment = NULL;
    wxButton* m_EEList_Button_SelectAll = NULL;
    wxButton* m_EEList_Button_SelectNone = NULL;



    //Switch buttons
    wxButton* m_switch_Button_Reestr = NULL;
    wxButton* m_switch_Button_DateDiff = NULL;
    wxButton* m_switch_Button_Backup = NULL;
    wxButton* m_switch_Button_MailMerge = NULL;
    wxButton* m_switch_Button_VacSchedule = NULL;

    wxCheckBox* m_Checkbox_DebugMode = NULL;



    wxTimer m_timer;


private:



    wxDECLARE_EVENT_TABLE();
};

wxBEGIN_EVENT_TABLE(MyFrame, wxFrame)

EVT_CHECKBOX(wx_chkBox_All, MyFrame::OnCheckBox)

EVT_BUTTON(wx_reestr_StartSearchingReestrs, MyFrame::OnReestrButtonPressed)
EVT_BUTTON(wx_DateDiff_ExecuteDateDifference, MyFrame::OnDateDiffButtonPressed)
EVT_BUTTON(wx_Backup_ExecuteBackup, MyFrame::OnBackupButtonPressed)

EVT_BUTTON(wx_reestr_Search_EmployeeName, MyFrame::OnReestrButtonPressed)
EVT_MENU(wx_reestr_OpenReestrOrders, MyFrame::OnOpenReestrOrders)
EVT_MENU(wx_reestr_OpenReestrDocs, MyFrame::OnOpenReestrDocs)
EVT_MENU(wx_reestr_OpenReestrAgreements, MyFrame::OnOpenReestrAgreements)

EVT_BUTTON(wx_reestr_StartMassReminder, MyFrame::OnReestrButtonMassReminder)

EVT_BUTTON(wx_Backup_AddFolder, MyFrame::OnBackupButtonAddFolder)
EVT_BUTTON(wx_Backup_AddFolderAlt, MyFrame::OnBackupButtonAddFolderAlt)
EVT_BUTTON(wx_Backup_AddFiles, MyFrame::OnBackupButtonAddFiles)
EVT_BUTTON(wx_Backup_SaveFilesToINI, MyFrame::OnBackupButtonSaveFilesToINI)

EVT_BUTTON(wx_MailMerge_ExecuteMailMerge, MyFrame::OnMailMergeButtonPressed)
EVT_BUTTON(wx_MailMerge_AddAttachmentFolder, MyFrame::OnMailMergeAddAttachmentFolder)
EVT_BUTTON(wx_MailMerge_AddAttachmentFiles, MyFrame::OnMailMergeAddAttachmentFiles)

EVT_BUTTON(wx_EElist_AddEEBase, MyFrame::OnEElistAddBase)
EVT_BUTTON(wx_EElist_SelectAll, MyFrame::OnEElistSelectAll)
EVT_BUTTON(wx_EElist_SelectNone, MyFrame::OnEElistSelectNone)
EVT_CHECKLISTBOX(wx_EElist_Department, MyFrame::OnEElistDepartmentSelected)

EVT_MENU(wx_EElist_CreateExampleBase, MyFrame::OnEElistCreateExampleBase)




EVT_MENU(wx_VacSchedule_Menu_OpenBase, MyFrame::OnVacScheduleOpenEEBase)


EVT_CHECKBOX(wx_chkBox_All, MyFrame::OnCheckBox)

EVT_LISTBOX(wx_VacSchedule_EEList, MyFrame::OnVacScheduleEEListSelected)


EVT_BUTTON(wx_SwitchPanelsReestr, MyFrame::OnSwitchPanelsButton)
EVT_BUTTON(wx_SwitchPanelsDateDifference, MyFrame::OnSwitchPanelsButton)
EVT_BUTTON(wx_SwitchPanelsBackup, MyFrame::OnSwitchPanelsButton)
EVT_BUTTON(wx_SwitchPanelsMailMerge, MyFrame::OnSwitchPanelsButton)
EVT_BUTTON(wx_SwitchPanelsVacSchedule, MyFrame::OnSwitchPanelsButton)


EVT_MENU(wxID_EXIT, MyFrame::OnExit)
EVT_MENU(wxID_ABOUT, MyFrame::OnAbout)

EVT_MENU(wx_OnLicence, MyFrame::OnLicence)
EVT_MENU(wx_OnProgramWebsite, MyFrame::OnProgramWebsite)



wxEND_EVENT_TABLE()







#include <vector>
using std::vector;

class WorkXPDatePickerPair
{
public:
    int iIDDatePickerFrom = -1;
    int iIDDatePickerTo = -1;
    wxDatePickerCtrl* DatePickerFrom = NULL;
    wxDatePickerCtrl* DatePickerTo = NULL;
};



class Panel
{
public:
    int iID = 0;
    int iShowDebugText = 1;
    int iDebugTextPos = 700;
    int iScrollBarHeight = 19;
    wxButton* SwitchButton = NULL;
    vector<wxControl*> Elements;
    vector<WorkXPDatePickerPair*> WorkExpDatePickers;
};

vector<Panel> g_PanelsArray;
vector<wxDatePickerCtrl*> g_DatePickersArray;



#import "lib\mso.dll" named_guids 
#import "lib\msoutl.olb"  rename("GetOrganizer", "GetOrganizerAddressEntry")
//#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52" rename("EOF","EOFile") rename("GetOrganizer", "GetOrganizerAddressEntry1")
using namespace Office;


//#import "libid:00062FFF-0000-0000-C000-000000000046" rename("EOF","EOFile") rename("GetOrganizer", "GetOrganizerAddressEntry2")







std::vector<std::string> SplitString(std::string str, std::string delimeter)
{
    std::vector<std::string> splittedStrings = {};
    size_t pos = 0;

    while ((pos = str.find(delimeter)) != std::string::npos)
    {
        std::string token = str.substr(0, pos);
        if (token.length() > 0)
            splittedStrings.push_back(token);
        str.erase(0, pos + delimeter.length());
    }

    if (str.length() > 0)
        splittedStrings.push_back(str);
    return splittedStrings;
}

std::vector<std::wstring> SplitWString(std::wstring str, std::wstring delimeter)
{
    std::vector<std::wstring> splittedStrings = {};
    size_t pos = 0;

    while ((pos = str.find(delimeter)) != std::wstring::npos)
    {
        std::wstring token = str.substr(0, pos);
        if (token.length() > 0)
            splittedStrings.push_back(token);
        str.erase(0, pos + delimeter.length());
    }

    if (str.length() > 0)
        splittedStrings.push_back(str);
    return splittedStrings;
}






#include <codecvt>
#include <climits>
#include <cstdint>
#include <cstdlib>
#include <cstring>
#include <iostream>
#include <sstream>
#include <string>
#include <vector>


using std::uint8_t;


#include "SimpleIni.h"
#include <windows.h>      // For common windows data types and function headers
#define STRICT_TYPED_ITEMIDS
#include <shlobj.h>
#include <objbase.h>      // For COM headers
#include <shobjidl.h>     // for IFileDialogEvents and IFileDialogControlEvents
#include <shlwapi.h>
#include <knownfolders.h> // for KnownFolder APIs/datatypes/function headers
#include <propvarutil.h>  // for PROPVAR-related functions
#include <propkey.h>      // for the Property key APIs/datatypes
#include <propidl.h>      // for the Property System APIs
#include <strsafe.h>      // for StringCchPrintfW
#include <shtypes.h>      // for COMDLG_FILTERSPEC
#include <new>

#include <filesystem>
#include <windows.h>
#include <shellapi.h>
#include <crtdbg.h>
#include <fstream>

///////
#include "OpenXLSX.hpp"
#include "pugixml.hpp"

#include <iomanip>
#include <iostream>
////////////
#include <boost/locale.hpp>
#include <boost/algorithm/string.hpp>
#include <boost/algorithm/string/split.hpp>
#include <boost/algorithm/string_regex.hpp>


#define Iter (*it)
#define Iter2 (*it2)
#define Iter3 (*it3)

using namespace std;
using namespace OpenXLSX;

#include "WorkerClass.h"

int g_VersionMajor = 1;
int g_VersionMinor = 1;

int g_CurrentPanel = -1000;
int g_FileUniqueCount = 0;
int g_GlobalThresholdUse = 0;
int g_GlobalThresholdValue = 0;
int g_DebugMode = 1;
int g_isBusy = 0;

int g_MailMerge_SendOnlyOneLetter = 0;
int g_MailMerge_LookforAttachments = 0;
int g_MailMerge_ShowLetters = 0;
int g_MailMerge_LookOnlyByAttachments = 0;
int g_MailMerge_AddStaticAttachments = 0;
int g_Backup_OpenFolder = 0;
int g_UseReestrSearch = 0;
int g_UseVacSchedule = 0;

std::wstring g_ReestrOrdersPath{};
std::wstring g_ReestrDocsPath{};
std::wstring g_ReestrAgreementsPath{};
std::wstring g_BackupFolderPath{};
std::wstring g_BackupFolderPathAlt{};
std::wstring g_MailMergeEEBasePath{};
std::wstring g_MailMergeAttachmentFolderPath{};

std::wstring g_VacScheduleEEBasePath{};

std::wstring g_INIPath = L"Options.ini";


int g_FilesLoaded = 0;
int g_ReestrDocsPathLoaded = 0;
int g_ReestrOrdersPathLoaded = 0;
int g_ReestrAgreementsPathLoaded = 0;
int g_MailMergeEEBasePathLoaded = 0;


#define t_File_Reestr_Docs 0
#define t_File_Reestr_Orders 1
#define t_File_Reestr_Agreements 2

#define t_File_Backup_Files 10
#define t_File_Backup_Folder 11
#define t_File_Backup_FolderAlt 12

#define t_File_MailMerge_EEBase 21
#define t_File_MailMerge_AttachmentFolder 22
#define t_File_MailMerge_AttachmentFiles 23

#define t_File_EElist_BaseToSave 30

#define t_File_VacSchedule_EEBase 40



int t_Panel_Max = -1;
#define t_Panel_Reestr 0
#define t_Panel_DateDifference 1
#define t_Panel_Backup 2
#define t_Panel_MailMerge 3
#define t_Panel_VacSchedule 4


#define t_Sheet_Reestr_Orders 0
#define t_Sheet_Reestr_Overtime 1
#define t_Sheet_Reestr_Docs 2
#define t_Sheet_Reestr_Agreements 3
#define t_Sheet_Reestr_Vacations 4

MyFrame* Frame_Main = NULL;
wxScrolledWindow* g_panel_Reestr = NULL;

int ReestrSearch_iFoundName = 0;
int ReestrSearch_iFoundDocuments = 0;
wstring ReestrSearch_wOutputResult{};
wstring ReestrSearch_wOutputDebug{};

HWND g_hWnd;

int g_FrameWidth = 1400;
int g_FrameWidthMinusScrollBar = g_FrameWidth - 50;
int g_FrameHeight = 1000;

void f_Add_Menus()
{


    wxMenu* fileMenu = new wxMenu;

    if (g_UseReestrSearch)
    {
        fileMenu->Append(wx_reestr_OpenReestrOrders, "Открыть реестр приказов...", "Открыть реестр приказов");
        fileMenu->Append(wx_reestr_OpenReestrDocs, "Открыть реестр получения документов...", "Открыть реестр получения документов");
        fileMenu->Append(wx_reestr_OpenReestrAgreements, "Открыть реестр доп. соглашений...", "Открыть реестр доп. соглашений");
        fileMenu->AppendSeparator();
    }

    fileMenu->Append(wx_EElist_CreateExampleBase, "Выгрузить шаблон базы работников");
    fileMenu->AppendSeparator();
    fileMenu->Append(wx_VacSchedule_Menu_OpenBase, "Загрузить базу для Графика отпусков");
    fileMenu->AppendSeparator();


    fileMenu->Append(wxID_EXIT, "Выход\tAlt-X", "Выход из программы");

    wxMenu* const helpMenu = new wxMenu;
    helpMenu->Append(wx_OnLicence, "Лицензия");
    helpMenu->Append(wx_OnProgramWebsite, "Посетить сайт программы");
    helpMenu->Append(wxID_ABOUT, "О программе");

    wxMenuBar* menuBar = new wxMenuBar();
    menuBar->Append(fileMenu, "Файл");
    menuBar->Append(helpMenu, "Помощь");
    Frame_Main->SetMenuBar(menuBar);
}





void f_PanelSetVisibility(int iPanel, int iShow)
{

    for (std::vector<Panel>::iterator it = g_PanelsArray.begin(); it != g_PanelsArray.end(); ++it)
    {
        if ((*it).iID == iPanel)
        {
            for (std::vector<wxControl*>::iterator Element = (*it).Elements.begin(); Element != (*it).Elements.end(); ++Element) { (*Element)->Show(iShow); }
        }
    }
}

void f_Panel_HideAllExcept(int iRequest)
{

    if (g_CurrentPanel == iRequest)
    {
        return;
    }
    else {
        g_CurrentPanel = iRequest;
    }

    CSimpleIniA ini;
    ini.SetUnicode();
    auto errVal = ini.LoadFile(g_INIPath.c_str());
    if (errVal != SI_Error::SI_FILE)
    {
        ini.SetLongValue("Settings", "iLastPanel", g_CurrentPanel);
        ini.SaveFile(g_INIPath.c_str());
    }

    for (std::vector<Panel>::iterator it = g_PanelsArray.begin(); it != g_PanelsArray.end(); ++it)
    {

        if ((*it).iID == iRequest)
        {
            //(*it).SwitchButton->SetBackgroundColour(*wxGREEN);
            //f_PanelSetVisibility((*it).iID, 1);
        }
        else {
            (*it).SwitchButton->SetBackgroundColour(*wxWHITE);
            f_PanelSetVisibility((*it).iID, 0);
        }

    }
    
    for (std::vector<Panel>::iterator it = g_PanelsArray.begin(); it != g_PanelsArray.end(); ++it)
    {

        if ((*it).iID == iRequest)
        {
            (*it).SwitchButton->SetBackgroundColour(*wxGREEN);
            f_PanelSetVisibility((*it).iID, 1);
            Frame_Main->m_text->Show((*it).iShowDebugText);
            Frame_Main->m_text->SetPosition(wxPoint(0, Iter.iDebugTextPos));
            g_panel_Reestr->SetScrollbars(20, Iter.iScrollBarHeight, 50, 50);

        }

    }



}


std::string f_wString_to_string(wstring wIn)
{
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> convert;
    std::string utf8String = convert.to_bytes(wIn);
    return utf8String;

}

std::wstring f_string_to_wString(std::string sIn)
{
    std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> convert;
    std::wstring utf16String = convert.from_bytes(sIn);
    return utf16String;
}


void f_MailMerge_ReplaceFields(wstring* w_LetterBody, wstring* w_LetterSubject, Worker* ChosenWorker)
{
    if (w_LetterSubject == NULL)
    {
        for (auto const& [key, val] : ChosenWorker->MergeFieldsMap)
        {
            boost::ireplace_all((*w_LetterBody), key, val);
        }

    }else {
        for (auto const& [key, val] : ChosenWorker->MergeFieldsMap)
        {
            boost::ireplace_all((*w_LetterBody), key, val);
            boost::ireplace_all((*w_LetterSubject), key, val);
        }

    }



}






bool f_SetINIValue(string sSection, string sKey, wstring wValue)
{
    CSimpleIniA ini;
    string sValue{};
    ini.SetUnicode();
    auto errVal = ini.LoadFile(g_INIPath.c_str());
    if (errVal != SI_Error::SI_FILE)
    {
        sValue = f_wString_to_string(wValue);
        ini.SetValue(sSection.c_str(), sKey.c_str(), sValue.c_str());
        ini.SaveFile(g_INIPath.c_str());
        return true;
    }

    return false;
}

bool f_SetINILong(string sSection, string sKey, int iValue)
{
    CSimpleIniA ini;
    ini.SetUnicode();
    auto errVal = ini.LoadFile(g_INIPath.c_str());
    if (errVal != SI_Error::SI_FILE)
    {
        ini.SetLongValue(sSection.c_str(), sKey.c_str(), iValue);
        ini.SaveFile(g_INIPath.c_str());
        return true;
    }

    return false;
}



//std::wstring widen(std::string const& in)
//{
//    std::wstring out{};
//
//    if (in.length() > 0)
//    {
//        // Calculate target buffer size (not including the zero terminator).
//        int len = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
//            in.c_str(), in.size(), NULL, 0);
//        if (len == 0)
//        {
//            throw std::runtime_error("Invalid character sequence.");
//        }
//
//        out.resize(len);
//        // No error checking. We already know, that the conversion will succeed.
//        MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
//            in.c_str(), in.size(), &out[0], out.size());
//        // Use out.data() in place of &out[0] for C++17
//    }
//
//    return out;
//}

string f_to_UTF8(std::string codepage_str)
{

    int size = MultiByteToWideChar(CP_ACP, MB_COMPOSITE, codepage_str.c_str(), codepage_str.length(), nullptr, 0);
    std::wstring utf16_str(size, '\0');
    MultiByteToWideChar(CP_ACP, MB_COMPOSITE, codepage_str.c_str(), codepage_str.length(), &utf16_str[0], size);

    int utf8_size = WideCharToMultiByte(CP_UTF8, 0, utf16_str.c_str(), utf16_str.length(), nullptr, 0, nullptr, nullptr);
    std::string utf8_str(utf8_size, '\0');
    WideCharToMultiByte(CP_UTF8, 0, utf16_str.c_str(), utf16_str.length(), &utf8_str[0], utf8_size, nullptr, nullptr);
    return utf8_str;
}


bool f_Does_file_exist(std::wstring p)
{
    return (std::filesystem::exists(p));
}


wstring f_GetWString_Version()
{

    //int IVersionTemp = 

    wstring w_TempName{};
    w_TempName += to_wstring(g_VersionMajor) + L".";

    //if (g_VersionMajor < 10)
    //{
    //    w_TempName += L"0";
    //}


    w_TempName += to_wstring(g_VersionMinor);


    return w_TempName;

}

void f_SetBusy(int iBusy)
{
    g_isBusy = iBusy;

    if (iBusy)
    {
        Frame_Main->SetTitle(L"HR Helper v" + f_GetWString_Version() + L" (подождите)");
    }
    else {
        Frame_Main->SetTitle(L"HR Helper v" + f_GetWString_Version());
    }
}

bool f_GetBusy()
{
    return g_isBusy;
}






#include "ExcelWork.h"
#include "OpenFileDialog.h"

#include "OnProgramStart.h"
#include "p_CreateElements.h"



void MyFrame::OnOpenReestrAgreements(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening new reestr agreements file\n");
    f_Open_File_Dialog(t_File_Reestr_Agreements);
}


void MyFrame::OnOpenReestrOrders(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening new reestr orders file\n");
    f_Open_File_Dialog(t_File_Reestr_Orders);
}

void MyFrame::OnOpenReestrDocs(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening base file\n");
    f_Open_File_Dialog(t_File_Reestr_Docs);
}

void MyFrame::OnEElistCreateExampleBase(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Creating example Base\n");
    f_Save_File_Dialog(t_File_EElist_BaseToSave);
}

void MyFrame::OnVacScheduleOpenEEBase(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening EE Vac schedule file\n");
    f_Open_File_Dialog(t_File_VacSchedule_EEBase);
}



void MyFrame::OnDateDiffButtonPressed(wxCommandEvent& WXUNUSED(event))
{
    int iDayFrom = Frame_Main->m_dateDiff_DatePickerFrom->GetValue().GetDay();
    int iMonthFrom = Frame_Main->m_dateDiff_DatePickerFrom->GetValue().GetMonth() + 1;
    int iYearFrom = Frame_Main->m_dateDiff_DatePickerFrom->GetValue().GetYear();

    int iDayTo = Frame_Main->m_dateDiff_DatePickerTo->GetValue().GetDay();
    int iMonthTo = Frame_Main->m_dateDiff_DatePickerTo->GetValue().GetMonth() + 1;
    int iYearTo = Frame_Main->m_dateDiff_DatePickerTo->GetValue().GetYear();


   int iDateFrom = DMYToExcelSerialDate(iDayFrom, iMonthFrom, iYearFrom);
    int iDateTo = DMYToExcelSerialDate(iDayTo, iMonthTo, iYearTo);
    int iDifference = iDateTo - iDateFrom + 1;

    m_text->AppendText("\nDays difference " + to_string(iDifference));
    //doBasicDemo();
}


#include "p_MailMerge.h"
#include "p_Reestr.h"
#include "p_Backup.h"
#include "p_VacSchedule.h"




void MyFrame::OnEElistAddBase(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("Opening EE Base for MailMerge \n");
    f_Open_File_Dialog(t_File_MailMerge_EEBase);
}

void MyFrame::OnEElistSelectAll(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("EE list select All \n");


    int iPos = 0, iCount = -1;


    iCount = m_EEList->GetCount();

    if (iCount<=0)
    {
        m_text->AppendText("\nБаза не загружена");
    }


    while (iPos < iCount)
    {
        m_EEList->Check(iPos, 1);
        iPos++;
    }



}

void MyFrame::OnEElistSelectNone(wxCommandEvent& WXUNUSED(event))
{
    m_text->AppendText("EE list select None \n");

    int iValue = -1, iPos = 0, iCount = -1;

    wxArrayInt CheckedItems;
    m_EEList->GetCheckedItems(CheckedItems);
    iCount = CheckedItems.GetCount();

    

    if (iCount <= 0)
    {
        m_text->AppendText("\nРаботники и так не указаны.");
        return;
    }

    while (iPos < iCount)
    {
        iValue = CheckedItems[iPos];
        m_EEList->Check(iValue, 0);
        iPos++;
    }
}


void MyFrame::OnEElistDepartmentSelected(wxCommandEvent& event)
{

    int iChecked = Frame_Main->m_EEListDepartment->IsChecked(event.GetInt());
    wstring w_Chosen_Department = event.GetString();

    if (iChecked)
    {
        Frame_Main->m_text->AppendText("\nВыбрано подразделение::");
    }
    else {
        Frame_Main->m_text->AppendText("\nУбрано подразделение::");
    }

    Frame_Main->m_text->AppendText(w_Chosen_Department);



    int iIndexFound = -1;

    for (std::vector<Worker*>::iterator it = g_WorkerArray.begin(); it != g_WorkerArray.end(); ++it)
    {
        Frame_Main->m_text->AppendText("\nCOMPARING::" + w_Chosen_Department + " with " + Iter->EEDepartment);

        if (boost::iequals(w_Chosen_Department, Iter->EEDepartment))
        {
            Frame_Main->m_text->AppendText("Found match");
            iIndexFound = Frame_Main->m_EEList->FindString(Iter->EENameFull);
            Frame_Main->m_text->AppendText("iIndexFound is " + to_string(iIndexFound));

            m_EEList->Check(iIndexFound, iChecked);
        }
    }



}





void MyFrame::OnCheckBox(wxCommandEvent& event)
{
    m_text->AppendText("\nCHECKED");

    if (event.GetEventObject() == m_MailMerge_Checkbox_LookforAttachments)
    {
        g_MailMerge_LookforAttachments = event.IsChecked();
        f_SetINILong("Settings", "bMailMerge_LookForAttachments", g_MailMerge_LookforAttachments);
    }
    else if (event.GetEventObject() == m_MailMerge_Checkbox_ShowLetters)
    {
        g_MailMerge_ShowLetters = event.IsChecked();
        f_SetINILong("Settings", "bMailMerge_ShowLetters", g_MailMerge_ShowLetters);
    }
    else if (event.GetEventObject() == m_Checkbox_DebugMode)
    {
        g_DebugMode = event.IsChecked();
        f_SetINILong("Settings", "b_DebugMode", g_DebugMode);

    }else if (event.GetEventObject() == m_Backup_Checkbox_OpenFolder)
    {
        g_Backup_OpenFolder = event.IsChecked();
        f_SetINILong("Settings", "b_Backup_OpenFolder", g_Backup_OpenFolder);

    }else if (event.GetEventObject() == m_MailMerge_Checkbox_LookOnlyByAttachments)
    {
        g_MailMerge_LookOnlyByAttachments = event.IsChecked();
        f_SetINILong("Settings", "bMailMerge_LookOnlyByAttachments", g_MailMerge_LookOnlyByAttachments);
    }else if (event.GetEventObject() == m_MailMerge_Checkbox_AddStaticAttachments)
    {
        g_MailMerge_AddStaticAttachments = event.IsChecked();
        f_SetINILong("Settings", "bMailMerge_AddStaticAttachments", g_MailMerge_AddStaticAttachments);
    }
    else if (event.GetEventObject() == m_MailMerge_Checkbox_SendOnlyOneLetter)
    {
        g_MailMerge_SendOnlyOneLetter = event.IsChecked();
        f_SetINILong("Settings", "bMailMerge_SendOnlyOneLetter", g_MailMerge_SendOnlyOneLetter);
    }

}


void MyFrame::OnVacScheduleEEListSelected(wxCommandEvent& event)
{
    Frame_Main->m_text->AppendText("\nВыбран::");
    wstring w_Chosen_WorkerName = event.GetString();
    Frame_Main->m_text->AppendText(w_Chosen_WorkerName);
    f_VacSchedule_Run_Selection_For(w_Chosen_WorkerName);
}



int MyApp::OnExit()
{


    //Frame_Main->DestroyChildren();

    //for (std::vector<wxDatePickerCtrl*>::iterator it = g_DatePickersArray.begin(); it != g_DatePickersArray.end(); ++it)
    //{

    //}


    return 1;
}




void MyFrame::OnAbout(wxCommandEvent& WXUNUSED(event))
{

    wstring w_text = L"HR Helper. Автор::praudmur. Версия>>" + f_GetWString_Version();
    w_text += L"\n \n Благодарности:";
    w_text += L"\n -troldal за OpenXLSX";
    w_text += L"\n -Разработчикам wxWidgets за wxWidgets";
    w_text += L"\n -brofield за simpleini";

    wxMessageBox
    (
        w_text,
        "О программе",
        wxOK | wxICON_INFORMATION,
        this
    );
}


void MyFrame::OnProgramWebsite(wxCommandEvent& WXUNUSED(event))
{
    int argc;
    char* argv[4000];

    int i = 1;
    string url = "https://github.com/praudmur/HR-Helper/releases";

    for (; i < argc; i++) {
        url = url + string(argv[i]);
        if (i != argc - 1) url = url + string("+");
    }

    //cout << url << endl;
    string op = string("start ").append(url);
    system(op.c_str());

}


void MyFrame::OnLicence(wxCommandEvent& WXUNUSED(event))
{

    wstring w_text = L"BSD 3-Clause License\n";
    w_text += L"Copyright(c) 2021, praudmur\n";
    w_text += L"All rights reserved.\n";

    w_text += L"Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:\n";
    w_text += L"1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.\n";
    w_text += L"2. Redistributions in binary form must reproduce the above copyright notice,this list of conditions and the following disclaimer in the documentationand /or other materials provided with the distribution.\n";
    w_text += L"3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.\n";


    w_text += L"THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS <<AS IS>>\n";
    w_text += L"AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE\n";
    w_text += L"IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE\n";
    w_text += L"DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE\n";
    w_text += L"FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL\n";
    w_text += L"DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR\n";
    w_text += L"SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER\n";
    w_text += L"CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,\n";
    w_text += L"OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE\n";
    w_text += L"OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.\n";


    wxMessageBox
    (
        w_text,
        "Лицензия",
        wxOK | wxICON_INFORMATION,
        this
    );
}


void MyFrame::OnSwitchPanelsButton(wxCommandEvent& event)
{

    if (event.GetEventObject() == Frame_Main->m_switch_Button_DateDiff)
    {
        f_Panel_HideAllExcept(t_Panel_DateDifference);
    }
    else if (event.GetEventObject() == Frame_Main->m_switch_Button_Reestr)
    {
        f_Panel_HideAllExcept(t_Panel_Reestr);
        
    }else if (event.GetEventObject() == Frame_Main->m_switch_Button_Backup)
    {
        f_Panel_HideAllExcept(t_Panel_Backup);
    }else if (event.GetEventObject() == Frame_Main->m_switch_Button_MailMerge)
    {
        f_Panel_HideAllExcept(t_Panel_MailMerge);
    }else if (event.GetEventObject() == Frame_Main->m_switch_Button_VacSchedule)
    {
        f_Panel_HideAllExcept(t_Panel_VacSchedule);
    }
}



#include <time.h>
#include <iostream>
#include <sstream>
#include <algorithm>

using namespace std;

#include <regex>
#include <string>



void MyFrame::OnMouseDoubleClick(wxMouseEvent& event)
{

    //if (event.GetEventType() == wxEVT_LEFT_DOWN)

    Frame_Main->m_text->AppendText("DOUBLE CLICK");

    if (event.GetEventObject() == Frame_Main->m_Backup_Status_BackupFolderPath)
    {
        if ( f_Does_file_exist(g_BackupFolderPath))
            ShellExecute(NULL, _T("explore"), g_BackupFolderPath.c_str(), NULL, NULL, SW_SHOW);


    }
    else if (event.GetEventObject() == Frame_Main->m_Backup_Status_BackupFolderPathAlt)
    {
        if (f_Does_file_exist(g_BackupFolderPathAlt))
            ShellExecute(NULL, _T("explore"), g_BackupFolderPathAlt.c_str(), NULL, NULL, SW_SHOW);
    }else if (event.GetEventObject() == Frame_Main->m_MailMerge_Status_AttachmentFolderPath)
    {
        if (f_Does_file_exist(g_MailMergeAttachmentFolderPath))
            ShellExecute(NULL, _T("explore"), g_MailMergeAttachmentFolderPath.c_str(), NULL, NULL, SW_SHOW);
    }


}


void MyFrame::OnKeyPress(wxKeyEvent& event)
{

    event.Skip();

    //m_reestr_Status_ReestrOrdersPath
    //m_reestr_Status_ReestrDocsPath
    //m_reestr_Status_ReestrAgreementsPath

    //if (wxGetKeyState(WXK_RETURN)) // ENTER
    //{



    //}


    if (wxGetKeyState(WXK_CONTROL_V) && wxGetKeyState(WXK_CONTROL))
    {

        for (std::vector<wxDatePickerCtrl*>::iterator it = g_DatePickersArray.begin(); it != g_DatePickersArray.end(); ++it)
        {
            if ((*it)->IsMouseInWindow())
            {

                OpenClipboard(0);
                char* pResult = (char*)GetClipboardData(CF_TEXT);

                if (pResult != NULL) {

                    wxDateTime DateChecker;
                    string str(pResult);

                    if (str.length() > 10) // input is not date
                    {
                        DateChecker.SetYear(-1000);
                        DateChecker.SetMonth((wxDateTime::Month)-1000);
                        DateChecker.SetDay(-1000);

                    }
                    else {

                        DateChecker.ParseDate(str);



                    }



                    if (DateChecker.IsValid())
                    {
                        (*it)->SetValue(DateChecker);
                        Frame_Main->m_text->AppendText("DATE VALID!!!!!!!!!/n");
                        break;
                    }
                    else {
                        Frame_Main->m_text->AppendText("DATE NOT VALID/n");
                    }
                }
                CloseClipboard();

            }
        }



    }

}
