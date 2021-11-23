using Microsoft.Office.Interop.Excel;
using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace SampleExportExcelFile
{
    public partial class Form1 : Form
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        //실행중 판단
        private bool IsRunExport = false;
        
        public Form1()
        {
            InitializeComponent();
        }

        //그리드 선택 안되게
        private void GridViewMaterial_Click(object sender, EventArgs e)
        {
            GridViewMaterial.ClearSelection();
        }

        private void ButtonLoad_Click(object sender, EventArgs e)
        {
            //그리드 리스트 초기화
            GridViewMaterial.Rows.Clear();
            
            //프로그레스바 정보 초기화
            ProgressBarExport.Value = 0;
            ProgressBarExport.Text = string.Empty;

            for (int i = 0; i < 7; i++)
            {
                var rnd = new Random(Guid.NewGuid().GetHashCode());
                var year = rnd.Next(DateTime.Now.Year, DateTime.Now.Year + 1);
                var month = rnd.Next(1, 13);
                var days = rnd.Next(1, DateTime.DaysInMonth(year, month) + 1);

                DateTime dateTime = new DateTime(year, month, days,
                    rnd.Next(0, 24), rnd.Next(0, 60), rnd.Next(0, 60), rnd.Next(0, 1000));

                string chars = "ABCDEFG";
                int sellerNumber = rnd.Next(1, chars.Length);
                string seller = $"Company{chars[sellerNumber]}";
                int count = rnd.Next(1, 100);

                DataGridViewRow AddRow = new DataGridViewRow();
                AddRow.CreateCells(GridViewMaterial);
                AddRow.Cells[0].Value = dateTime.ToString();
                AddRow.Cells[1].Value = seller;
                AddRow.Cells[2].Value = count.ToString();
                AddRow.Height = 25;
                GridViewMaterial.Rows.Add(AddRow);
            }

            GridViewMaterial.ClearSelection();
        }

        private delegate void CrossThreadSafetyProgressExport(int rowNum, int rowcount);
        public void SetProgressBarExport(int rowNum, int rowcount)
        {
            if (ProgressBarExport != null)
            {
                if (ProgressBarExport.InvokeRequired)
                {
                    try
                    {
                        if (ProgressBarExport != null)
                        {
                            ProgressBarExport.Invoke(new CrossThreadSafetyProgressExport(SetProgressBarExport), rowNum, rowcount);
                        }
                    }
                    finally { }
                }
                else
                {
                    try
                    {
                        if (rowcount != 0)
                        {
                            int percent = (int)(rowNum * 100 / rowcount);
                            ProgressBarExport.Value = percent;
                            ProgressBarExport.Text = $"{percent}% ({rowNum}/{rowcount})";
                        }
                        else
                        {
                            ProgressBarExport.Value = 100;
                            ProgressBarExport.Text = $"100% (0/0)";
                        }
                    }
                    finally { }
                }
            }
        }

        private void ButtonExport_Click(object sender, EventArgs e)
        {
            //실행 중일 때 간략 예외처리
            if (IsRunExport)
            {
                MessageBox.Show("Alerady Export", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            //데이터 없을 때 예외처리
            if (GridViewMaterial.Rows.Count <= 0)
            {
                MessageBox.Show("No Load Datqa", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string nowDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
            string savePath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string pathFilename = $"{savePath}\\excelexport_{nowDateTime}.xlsx";

            //사용자가 FileDialog를 사용하게 하려면
            {
                SaveFileDialog saveFile = new SaveFileDialog()
                {
                    Title = "Save Excel File",
                    FileName = $"excelexport_{nowDateTime}.xlsx",
                    DefaultExt = "xlsx",
                    Filter = "Xlsx files(*.xlsx)|*.xlsx"
                };
                // OK버튼을 눌렀을때의 동작
                if (saveFile.ShowDialog() == DialogResult.OK)
                {
                    // 경로와 파일명을 fileName에 저장
                    pathFilename = saveFile.FileName.ToString();
                }
                else
                {
                    return;
                }
            }

            //ex. 테두리를 위해 그리드 축 개수를 담아두고 (프로그레스바 카운트까지 미리)
            //프로그레스바 카운트 => Row로 진행
            int columnCount = GridViewMaterial.Columns.Count;
            int rowCount = GridViewMaterial.Rows.Count;

            //+프로그레스 바 초기화
            SetProgressBarExport(0, rowCount);

            //엑셀 파일 만들기 시작
            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
            excel.DisplayAlerts = false;

            //1. 워크시트 선택
            //처음에는 Sheet1로 1개 있음
            Worksheet worksheet = (Worksheet)workbook.Worksheets.Item["Sheet1"];
            //여러 시트를 하려면 인덱스를 추가해서 받아서 사용 (2번째 부터는)
            //workbook.Worksheets.Add(After: workbook.Worksheets[index - 1]);
            //Worksheet worksheet = workbook.Worksheets.Item[index];

            //2. 필요 시 시트 이름 변경
            worksheet.Name = "Sample Sheet Name";

            //3. 컬럼 별로 너비 변경
            Range ModRange = (Range)worksheet.Columns[1];
            ModRange.ColumnWidth = 30;
            ModRange = (Range)worksheet.Columns[2];
            ModRange.ColumnWidth = 30;
            //넘버포맷을 사용하면 뒤 컬럼부터는 숫자형식으로 적용
            ModRange.NumberFormat = "@";
            ModRange = (Range)worksheet.Columns[3];
            ModRange.ColumnWidth = 30;
            ModRange = (Range)worksheet.Columns[4];
            ModRange.ColumnWidth = 30;

            //4. 첫번째 줄 타이틀 생성 - 예쁘게 하기 위해
            //Range는 엑셀을 실행해서 참고하기 좋음 (첫줄이라 1라인)
            ModRange = (Range)worksheet.get_Range("A1", "C1");
            ModRange.Merge(true); //병합하고
            ModRange.Value = $"Export Excel File Title"; //이름 입력하고
            ModRange.Font.Size = 16; //폰트 키우고
            ModRange.Font.Bold = true; //Bold 주고
            ModRange.HorizontalAlignment = XlHAlign.xlHAlignLeft; //좌측 정렬
                                                                  //테두리 까지 끝
            ModRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlMedium, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);

            //5. 2번째 줄에는 리포트 기간 및 파일 설명 추가
            ModRange = (Range)worksheet.get_Range("A2", "C2");
            ModRange.Merge(true);
            //2번째 설명은 우측 정렬
            ModRange.HorizontalAlignment = XlHAlign.xlHAlignRight;

            //5. 헤드열 추가
            //cell은 1부터 row나 column은 일반적인 0부터라 차이가 있는 점 주의
            for (int i = 0; i < columnCount; i++)
            {
                ModRange = (Range)worksheet.Cells[3, 1 + i];
                ModRange.Value = GridViewMaterial.Columns[i].HeaderText;
                ModRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;

                //data 테두리
                ModRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                ModRange.Borders[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlMedium; //위 테두리
                if (i == 0) //시작 컬럼에서 왼쪽 테두리
                {
                    ModRange.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                }
                else if (i == (columnCount - 1)) //마지막 컬럼에서 우측 테두리
                {
                    ModRange.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                }
                //아래 2줄 얇은 테두리
                ModRange.Borders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlDouble;
                ModRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            }

            //6. 데이터 열 추가
            for (int i = 0; i < rowCount; i++)
            {
                //+프로그레스 바 진행 표시
                SetProgressBarExport(i+1, rowCount);

                for (int j = 0; j < columnCount; j++)
                {
                    //타이틀, 추가설명, 헤드, 0->1 때문에 i에 4를 더함
                    ModRange = (Range)worksheet.Cells[4 + i, 1 + j];
                    ModRange.Value = GridViewMaterial[j, i].Value == null ? string.Empty : GridViewMaterial[j, i].Value.ToString();

                    //data 테두리
                    ModRange.BorderAround2(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexAutomatic, XlColorIndex.xlColorIndexAutomatic);
                    if (j == 0) //시작 컬럼에서 왼쪽 테두리
                    {
                        ModRange.Borders[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlMedium;
                    }
                    else if (j == (columnCount - 1)) //마지막 컬럼에서 우측 테두리
                    {
                        ModRange.Borders[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlMedium;
                    }
                    if (i == (rowCount - 1)) //마지막 로우에서 우측 테두리
                    {
                        ModRange.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlMedium;
                        //결산 같은 마지막 줄 값이 존재하면 이걸 사용합니다.
                        //ModRange.Borders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlDouble;
                    }
                }
            }

            //7. 상단 고정필드 설정
            worksheet.Application.ActiveWindow.SplitRow = 1;
            worksheet.Application.ActiveWindow.FreezePanes = true;
            worksheet.Application.ActiveWindow.SplitRow = 2;
            worksheet.Application.ActiveWindow.FreezePanes = true;
            worksheet.Application.ActiveWindow.SplitRow = 3;
            worksheet.Application.ActiveWindow.FreezePanes = true;

            //8. 파일 저장 (앞선 SaveFileDialog로 만들어진 pathFilename 경로로 파일 저장
            workbook.SaveAs(Filename: pathFilename);
            workbook.Close();

            //9. 종료 안되는 excel 프로세스 수동 제거
            ProgressBarExport.Invoke((MethodInvoker)(() =>
            {
                ProgressBarExport.Invalidate();
            }));
            try
            {
                GetWindowThreadProcessId(new IntPtr(excel.Hwnd), out uint processId);
                excel.Quit();
                if (processId != 0)
                {
                    try
                    {
                        System.Diagnostics.Process excelProcess = System.Diagnostics.Process.GetProcessById((int)processId);
                        excelProcess.CloseMainWindow();
                        excelProcess.Refresh();
                        excelProcess.Kill();
                    }
                    catch { }
                }
            }
            catch { }

            //+프로그레스 바 완료
            SetProgressBarExport(rowCount, rowCount);

            MessageBox.Show("Complete Export.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            IsRunExport = false;
        }
    }
}
