using System;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using LINQtoCSV;
using SharedClassLibrary;
using SharedClassLibrary.Utilities;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Connections;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Ellipse.Connections;
using Debugger = SharedClassLibrary.Utilities.Debugger;


namespace BonoDeTopeados
{
    public partial class FormMainScreen : Form
    {
        private Thread _thread;
        private EllipseFunctions _eFunctions;
        private Application _excelApp;

        public FormMainScreen()
        {
            InitializeComponent();
            LoadSettings();
            cbPeriodMode.Items.Add("NORMAL");
            cbPeriodMode.Items.Add("MES CORRIDO");
            cbPeriodMode.Items.Add("MES FIJO");
            cbPeriodMode.SelectedIndex = 0;

            var currentDate = DateTime.Today.AddDays(-30);
            tbYear.Text = "" + currentDate.Year;
            tbPeriod.Text = "" + (currentDate.Month / 3);
        }
        public void LoadSettings()
        {
            var settings = new Settings();
            _eFunctions = new EllipseFunctions();

            //Setting of Configuration Options from Config File (or default)
            try
            {
                settings.LoadCustomSettings();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, SharedResources.Settings_Title, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            settings.SaveCustomSettings();
        }
        private void btnLoadEmployeeTurns_Click(object sender, EventArgs e)
        {
            var periodMode = "" + cbPeriodMode.SelectedItem;
            if (_thread != null && _thread.IsAlive) return;

            _thread = new Thread(() => LoadEmployeeTurns(periodMode));
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }

        public void LoadEmployeeTurns(string periodMode)
        {
            var ef = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo));
            var btYear = MyUtilities.ToInteger(tbYear.Text);
            var btPeriod = MyUtilities.ToInteger(tbPeriod.Text);
            DatePeriod.SetPeriod(btYear, btPeriod);
            try
            {
                //Abrir el archivo
                var openFileDialog1 = new OpenFileDialog
                {
                    Filter = @"Archivos CSV|*.csv",
                    FileName = @"LISTADO_EMP.csv",
                    Title = @"Instance Series Generator",
                    //InitialDirectory = @"C:\Data\Loaders\Parametros"
                };

                if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

                this.UseWaitCursor = true;
                var filePath = openFileDialog1.FileName;

                var inputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = ',',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };
                var cc = new CsvContext();
                var listadoEmpleadoFile = cc.Read<ListadoEmpleadoItem>(filePath, inputFileDescription);
                //
                Debugger.DebugginMode = false;

                EmployeeTurns prevEmp = null;
                var totalItems = listadoEmpleadoFile.Count();
                var currentItem = 0;
                foreach (var f in listadoEmpleadoFile)
                {

                    currentItem++;
                    Invoke((MethodInvoker)delegate ()
                    {
                        lblProgress.Text = currentItem + "/" + totalItems;
                    });

                    var curEmp = new EmployeeTurns(f, periodMode);
                    //Si no es del periodo
                    if (curEmp.Anho != btYear || curEmp.Periodo != btPeriod)
                        continue;

                    //Si es diferente  estas superintendencias
                    if (!(curEmp.DescSuperintendencia.Equals("PLANTAS DE CARBON") ||
                          curEmp.DescSuperintendencia.Equals("PLANTAS FLUJO Y CONT CAL CARB") ||
                          curEmp.DescSuperintendencia.Equals("SUPERINTENDENTE ASISTENTE OPM") ||
                          curEmp.DescSuperintendencia.Equals("FERROCARRIL") ||
                          curEmp.DescSuperintendencia.Equals("PUERTO") ||
                          curEmp.DescSuperintendencia.Equals("PLANEACION ANALISIS Y MEJORAM")))
                    {
                        continue;
                    }

                    if (prevEmp == null)
                        prevEmp = curEmp;
                    else if (curEmp.Equals(prevEmp, true))
                        prevEmp.SumTurns(curEmp);
                    else
                    {
                        //INGRESAR
                        Debugger.LogDebugging(prevEmp.Cedula + "\t" + prevEmp.Nombre + "\t" + prevEmp.Anho + "\t" + prevEmp.Periodo + "\t" + prevEmp.TurnoD +
                                              "\t" + prevEmp.TurnoD +
                                              "\t" + prevEmp.TurnoL +
                                              "\t" + prevEmp.TurnoI +
                                              "\t" + prevEmp.TurnoM +
                                              "\t" + prevEmp.TurnoN +
                                              "\t" + prevEmp.TurnoP +
                                              "\t" + prevEmp.TurnoT +
                                              "\t" + prevEmp.TurnoV +
                                              "\t" + prevEmp.TurnoOtro);
                        var sqlQuery = Queries.InsertEmployeeTurnType(prevEmp);
                        ef.GetQueryResult(sqlQuery);
                        prevEmp = curEmp;
                    }

                    
                    
                }

                
                MessageBox.Show("Proceso terminado");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.UseWaitCursor = false;
                ef.CloseConnection();
            }
        }

        public void LoadEmployeeCovidValue(string periodMode)
        {
            var ef = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo));
            var btYear = MyUtilities.ToInteger(tbYear.Text);
            var btPeriod = MyUtilities.ToInteger(tbPeriod.Text);
            DatePeriod.SetPeriod(btYear, btPeriod);
            try
            {
                //Abrir el archivo
                var openFileDialog1 = new OpenFileDialog
                {
                    Filter = @"Archivos CSV|*.csv",
                    FileName = @"Ausentismo.csv",
                    Title = @"Instance Series Generator",
                    //InitialDirectory = @"C:\Data\Loaders\Parametros"
                };

                if (openFileDialog1.ShowDialog() != DialogResult.OK) return;

                this.UseWaitCursor = true;
                var filePath = openFileDialog1.FileName;

                var inputFileDescription = new CsvFileDescription
                {
                    SeparatorChar = ',',
                    FirstLineHasColumnNames = true,
                    EnforceCsvColumnAttribute = true
                };
                var cc = new CsvContext();
                var listadoEmpleadoFile = cc.Read<AusentismoEmpleadoItem>(filePath, inputFileDescription);
                //
                Debugger.DebugginMode = true;

                EmployeeTurns prevEmp = null;
                var totalItems = listadoEmpleadoFile.Count();
                var currentItem = 0;
                foreach (var f in listadoEmpleadoFile)
                {
                    currentItem++;
                    Invoke((MethodInvoker)delegate ()
                    {
                        lblProgress.Text = currentItem + "/" + totalItems;
                    });

                    if (prevEmp == null)
                        prevEmp = GetEmployeeTurn(f.EmployeeId, btYear, btPeriod, periodMode);
                    if (prevEmp == null)
                        continue;


                    if (f.EmployeeId.Equals(prevEmp.Cedula))
                    {
                        var filePeriod = DatePeriod.GetPeriod(f.EmployeeId, prevEmp.CodDependencia, f.Anho, f.Mes, periodMode);
                        if (filePeriod.Year != btYear || filePeriod.Period != btPeriod)
                            continue;
                        prevEmp.TurnoOtro += f.Hr886;
                    }
                    else
                    {
                        //INGRESAR
                        Debugger.LogDebugging(prevEmp.Cedula + "\t" + prevEmp.Nombre + "\t" + prevEmp.Anho + "\t" + prevEmp.Periodo + "\t" +
                                              //"\t" + prevEmp.TurnoD +
                                              //"\t" + prevEmp.TurnoL +
                                              //"\t" + prevEmp.TurnoI +
                                              //"\t" + prevEmp.TurnoM +
                                              //"\t" + prevEmp.TurnoN +
                                              //"\t" + prevEmp.TurnoP +
                                              //"\t" + prevEmp.TurnoT +
                                              //"\t" + prevEmp.TurnoV +
                                              "\t" + prevEmp.TurnoOtro);


                        var sqlQuery = Queries.InsertEmployeeTurn886(prevEmp);
                        ef.ExecuteQuery(sqlQuery);


                        prevEmp = GetEmployeeTurn(f.EmployeeId, btYear, btPeriod, periodMode);
                        if (prevEmp != null)
                        {
                            var filePeriod = DatePeriod.GetPeriod(f.EmployeeId, prevEmp.CodDependencia, f.Anho, f.Mes, periodMode);
                            if (filePeriod.Year != btYear || filePeriod.Period != btPeriod)
                                continue;
                            prevEmp.TurnoOtro += f.Hr886;
                        }
                    }



                }

                if (prevEmp != null)
                {
                    var sqlQuery = Queries.InsertEmployeeTurn886(prevEmp);
                    ef.ExecuteQuery(sqlQuery);
                }
                MessageBox.Show("Proceso terminado");
            }
            catch (Exception ex)
            {
                
                MessageBox.Show(ex.Message);
            }
            finally
            {
                ef.CloseConnection();
                this.UseWaitCursor = false;
            }
        }

        public EmployeeTurns GetEmployeeTurn(string cedula, int anho, int period, string periodMode)
        {
            var dbConn = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo));
            var empTurn = new EmployeeTurns();

            var sqlQuery = Queries.GetEmployeeTurnType(cedula, anho, period);
            var dReader = dbConn.GetQueryResult(sqlQuery);
            if (dReader == null || dReader.IsClosed || !dReader.Read())
            {
                dbConn.CloseConnection();
                return null;
            }

            empTurn.Cedula = "" + dReader["CEDULA"].ToString().Trim();
            empTurn.Anho = MyUtilities.ToInteger("" + dReader["ANO"].ToString().Trim());
            empTurn.Periodo = MyUtilities.ToInteger("" + dReader["PERIODO"].ToString().Trim());
            empTurn.Nombre = "" + dReader["NOMBRE"].ToString().Trim();
            empTurn.Cargo = "" + dReader["CARGO"].ToString().Trim();
            empTurn.CodSuperintendencia = MyUtilities.ToInteger("" + dReader["COD_SUPERINTENDENCIA"].ToString().Trim());
            empTurn.DescDependencia = "" + dReader["SUPERINTENDENCIA"].ToString().Trim();
            empTurn.CodDependencia = MyUtilities.ToInteger("" + dReader["COD_DEPENDENCIA"].ToString().Trim());
            empTurn.DescDependencia = "" + dReader["DESC_DEPENDENCIA"].ToString().Trim();
            empTurn.Supervisor = "" + dReader["SUPERVISOR"].ToString().Trim();
            empTurn.NivelEmpleado = MyUtilities.ToInteger("" + dReader["NVL_EMP"].ToString().Trim());
            empTurn.NivelCargo = MyUtilities.ToInteger("" + dReader["NVL_CARGO"].ToString().Trim());
            empTurn.Rol = MyUtilities.ToInteger("" + dReader["ROL"].ToString().Trim());
            empTurn.Estado = "" + dReader["ESTADO"].ToString().Trim();
            empTurn.TurnoD = MyUtilities.ToInteger("" + dReader["TURNO_D"].ToString().Trim());
            empTurn.TurnoL = MyUtilities.ToInteger("" + dReader["TURNO_L"].ToString().Trim());
            empTurn.TurnoI = MyUtilities.ToInteger("" + dReader["TURNO_I"].ToString().Trim());
            empTurn.TurnoM = MyUtilities.ToInteger("" + dReader["TURNO_M"].ToString().Trim());
            empTurn.TurnoN = MyUtilities.ToInteger("" + dReader["TURNO_N"].ToString().Trim());
            empTurn.TurnoP = MyUtilities.ToInteger("" + dReader["TURNO_P"].ToString().Trim());
            empTurn.TurnoT = MyUtilities.ToInteger("" + dReader["TURNO_T"].ToString().Trim());
            empTurn.TurnoV = MyUtilities.ToInteger("" + dReader["TURNO_V"].ToString().Trim());
            //empTurn.TurnoOtro = MyUtilities.ToInteger("" + dReader["TURNO_OTRO"].ToString().Trim());

            dbConn.CloseConnection();
            return empTurn;
        }

        private void btnLoadEmployeeTurn886_Click(object sender, EventArgs e)
        {
            var periodMode = "" + cbPeriodMode.SelectedItem;
            if (_thread != null && _thread.IsAlive) return;

            _thread = new Thread(() => LoadEmployeeCovidValue(periodMode));
            _thread.SetApartmentState(ApartmentState.STA);
            _thread.Start();
        }
    }
}