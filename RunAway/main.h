//---------------------------------------------------------------------------

#ifndef mainH
#define mainH
//---------------------------------------------------------------------------
#include <Classes.hpp>
#include <Controls.hpp>
#include <StdCtrls.hpp>
#include <Forms.hpp>
#include <Dialogs.hpp>
#include <Grids.hpp>
#include <FileCtrl.hpp>
//---------------------------------------------------------------------------
// Для работы с EXCEL
#include      <ComObj.hpp>
#include      <utilcls.h>
//---------------------------------------------------------------------------
class TVar
{
public:		// User declarations
	__fastcall TVar();
	AnsiString NameVar[100];
	int TypeVar[100];
	int Dimension[100];
	int NumVals[100];
	int MaxVals;
	int NumDays;
	int MaxDays;
	int IntValue[100][100];
	TDateTime DateValue[100][100];
};

//---------------------------------------------------------------------------
class TForm1 : public TForm
{
__published:	// IDE-managed Components
	TGroupBox *GroupBox1;
	TButton *Button1;
	TButton *Button2;
	TMemo *Memo1;
	TEdit *Edit1;
	TOpenDialog *OpenDialog1;
	TButton *Button3;
	TFileListBox *FileListBox1;
	TRadioButton *RadioButton1;
	TRadioButton *RadioButton2;
	TRadioButton *RadioButton3;
	TRadioButton *RadioButton4;
	TRadioButton *RadioButton5;
	TCheckBox *CheckBox1;
	TButton *Button4;
	TButton *Button5;
	void __fastcall Button2Click(TObject *Sender);
	void __fastcall Button3Click(TObject *Sender);
	void __fastcall Button1Click(TObject *Sender);

	void __fastcall ReadSourceFile(AnsiString FileName, int NDay);
	void __fastcall ShowReport(int NDay);
	void __fastcall ExportReport(int NDay);
	void __fastcall FileListBox1Click(TObject *Sender);
	void __fastcall RadioButton3Click(TObject *Sender);
	void __fastcall RadioButton4Click(TObject *Sender);
	void __fastcall RadioButton2Click(TObject *Sender);
	void __fastcall RadioButton1Click(TObject *Sender);
	void __fastcall RadioButton5Click(TObject *Sender);
	void __fastcall CheckBox1Click(TObject *Sender);
	void __fastcall Button4Click(TObject *Sender);
	void __fastcall Button5Click(TObject *Sender);
	private:	// User declarations
public:		// User declarations
	AnsiString SourcesDirectory;
	AnsiString ReportsDirectory;
	AnsiString SourceRunAwayKFileName;
	AnsiString SourceRunAwayTFileName;
	AnsiString SourceRunAwayPFileName;
	AnsiString SourceRunAwayKPSFileName;
	AnsiString DefaultSourceRunAwayFileName;
	AnsiString FullSourceRunAwayKFileName;
	AnsiString FullSourceRunAwayTFileName;
	AnsiString FullSourceRunAwayPFileName;
	AnsiString FullSourceRunAwayKPSFileName;
	AnsiString SourceRunAwayFileName;

	TVar Vars[100];

	Variant  vVarApp,vVarBooks,vVarBook,vVarSheets,
				 vVarSheet,vVarCells,vVarCell;
	bool fStart;
	int MaxVars;
	int NumVars;
	int FiveMinMode;
	int Mode;

	__fastcall TForm1(TComponent* Owner);
};
//---------------------------------------------------------------------------
extern PACKAGE TForm1 *Form1;
//---------------------------------------------------------------------------
#endif
