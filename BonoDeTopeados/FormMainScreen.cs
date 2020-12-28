using System;
using System.Threading;
using System.Windows.Forms;
using LINQtoCSV;
using SharedClassLibrary.Connections.Oracle;
using SharedClassLibrary.Ellipse;
using SharedClassLibrary.Ellipse.Connections;
using SharedClassLibrary.Utilities;
using Debugger = SharedClassLibrary.Utilities.Debugger;


namespace BonoDeTopeados
{
    public partial class FormMainScreen : Form
    {
        private Thread _thread;


        public FormMainScreen()
        {
            InitializeComponent();

            cbPeriodMode.Items.Add("NORMAL");
            cbPeriodMode.Items.Add("MES CORRIDO");
            cbPeriodMode.Items.Add("MES FIJO");
            cbPeriodMode.SelectedIndex = 0;
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

                var ef = new EllipseFunctions();
                ef.SetDBSettings(Environments.SigcorProductivo);

                foreach (var f in listadoEmpleadoFile)
                {
                    if(prevEmp == null)
                        prevEmp = new EmployeeTurns(f, periodMode);
                    var curEmp = new EmployeeTurns(f, periodMode);

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

                    if (curEmp.Equals(prevEmp, true))
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
                this.UseWaitCursor = false;
                MessageBox.Show("Proceso terminado");
            }
            catch (Exception ex)
            {
                this.UseWaitCursor = false;
                MessageBox.Show(ex.Message);
            }
        }

        public void LoadEmployeeCovidValue(string periodMode)
        {
            var ef = new EllipseFunctions();
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
                Debugger.DebugginMode = false;

                EmployeeTurns prevEmp = null;

                ef.SetDBSettings(Environments.SigcorProductivo);

                foreach (var f in listadoEmpleadoFile)
                {
                    if (prevEmp == null)
                        prevEmp = GetEmployeeTurn(f.EmployeeId, f.Anho, f.Mes, periodMode);

                    if (prevEmp == null)
                        continue;
                    if (f.EmployeeId.Equals(prevEmp.Cedula) && prevEmp.Period(prevEmp.Cedula, prevEmp.CodDependencia, f.Mes, periodMode) == prevEmp.Periodo)
                        prevEmp.TurnoOtro += f.Hr886;
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
                        if (prevEmp.Periodo == 2 || prevEmp.Periodo == 3)
                        {
                            var sqlQuery = Queries.InsertEmployeeTurn886(prevEmp);
                            ef.GetQueryResult(sqlQuery);
                        }

                        prevEmp = GetEmployeeTurn(f.EmployeeId, f.Anho, f.Mes, periodMode);
                        if (prevEmp != null)
                            prevEmp.TurnoOtro += f.Hr886;
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
                ef.CloseConnection();
                this.UseWaitCursor = false;
            }
        }

        public EmployeeTurns GetEmployeeTurn(string cedula, int anho, int mes, string periodMode)
        {
            var ef = new OracleConnector(Environments.GetDatabaseItem(Environments.SigcorProductivo));
            var empTurn = new EmployeeTurns();

            var sqlQuery = Queries.GetEmployeeTurnType(cedula, anho);
            var dReader = ef.GetQueryResult(sqlQuery);
            if (dReader == null || dReader.IsClosed)
            {
                ef.CloseConnection();
                return null;
            }

            dReader.Read();
            empTurn.Cedula = "" + dReader["CEDULA"].ToString().Trim();
            empTurn.Anho = MyUtilities.ToInteger("" + dReader["ANO"].ToString().Trim());
            //empTurn.Periodo = MyUtilities.ToInteger("" + dReader["PERIODO"].ToString().Trim());
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

            empTurn.Periodo = empTurn.Period(empTurn.Cedula, empTurn.CodDependencia, mes, periodMode);
            ef.CloseConnection();
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