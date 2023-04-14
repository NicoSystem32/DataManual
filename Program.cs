using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using ClosedXML.Excel;
using Dapper;
using DocumentFormat.OpenXml.Spreadsheet;
using Oracle.ManagedDataAccess.Client;
using Programa;

namespace ExcelExample
{
    public class Program
    {
        public static void Main(string[] args)
        {

            int opcion;
            do
            {
                Console.WriteLine("Seleccione una opción:");
                Console.WriteLine("1. Asociar Ids a Excel");
                Console.WriteLine("2. Insertar Finiquitos Manuales a Tabla");
                Console.WriteLine("3. Asociar KG a tabla de demostración");
                Console.WriteLine("0. Salir");
                opcion = int.Parse(Console.ReadLine());

                switch (opcion)
                {
                    case 1:
                        // Lógica para asociar Ids a Excel
                        Console.WriteLine("Se van asociar los Ids de los DCD´s al excel");
                        Console.WriteLine("Cargando...");
                        for (int i = 0; i < 20; i++)
                        {
                            Console.Write("\u2588");
                            Thread.Sleep(500);
                        }
                        Console.WriteLine();
                        // Obtiene la ruta del archivo Excel
                        string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                        string solutionDirectory = Path.GetFullPath(Path.Combine(baseDirectory, @"..\..\..\"));
                        string filePath = Path.Combine(solutionDirectory, "Excel", "Manuales.xlsx");

                        // Conexión a la base de datos Oracle
                        string connectionString = "DATA SOURCE=192.168.100.92:1521/SIFFANT;PERSIST SECURITY INFO=True;USER ID=fondosweb; PASSWORD=PRUEBAS2015;";
                        string sql = "SELECT * FROM F_DOCCOMPROMISODESTINO";
                        string sqlEntidad = "SELECT * FROM F_ENTIDAD";


                        using (var connection = new OracleConnection(connectionString))
                        {
                            var results = connection.Query<DCDModel>(sql).ToList();
                            List<DCDModel> dcdList = results.Where(x => x.DCD_CONSECUTIVO != null).ToList();
                            var entidadResults = connection.Query<EntidadModel>(sqlEntidad).ToList();

                            // Itera a través de las filas y columnas del archivo Excel
                            try
                            {
                                using (var workbook = new XLWorkbook(filePath))
                                {
                                    var worksheet = workbook.Worksheet(1);

                                    int rows = worksheet.RowsUsed().Count();
                                    int columns = worksheet.ColumnsUsed().Count();

                                    // Columna donde se escribirá el valor IdDCD
                                    int idDcdColumn = 10;
                                    // Columna donde se escribirá el valor NIT Comercializadora
                                    int nitComercializadoraColumn = 13;
                                    // Columna donde se escribirá el valor EntidadCodigo Comercializadora
                                    int entidadCodeColumn = 14;
                                    // Columna donde se escribirá el valor CodProveedor
                                    int codProvCodeColumn = 15;
                                    // Columna donde se escribirá el valor Grupo Mercado
                                    int groupMerColumn = 16;

                                    // Encabezados de las nuevas columnas
                                    worksheet.Cell(1, idDcdColumn).Value = "IdDCD";
                                    worksheet.Cell(1, nitComercializadoraColumn).Value = "NIT";
                                    worksheet.Cell(1, entidadCodeColumn).Value = "CodEntidad";
                                    worksheet.Cell(1, codProvCodeColumn).Value = "CodProv";
                                    worksheet.Cell(1, groupMerColumn).Value = "GrupoMercado";

                                    for (int row = 2; row <= rows; row++)
                                    {
                                        // Obtiene el valor de la columna DOCUMENTO de la fila actual
                                        string documento = worksheet.Cell(row, 6).Value.ToString();

                                        // Obtiene la suma de las columnas KILPALMA y KILPALMISTE de la fila actual
                                        double kilpalma;
                                        if (!double.TryParse(worksheet.Cell(row, 7).Value.ToString(), out kilpalma))
                                        {
                                            kilpalma = 0;
                                        }

                                        double kilpalmsite;
                                        if (!double.TryParse(worksheet.Cell(row, 8).Value.ToString(), out kilpalmsite))
                                        {
                                            kilpalmsite = 0;
                                        }

                                        // Busca el registro correspondiente en la tabla F_DOCCOMPROMISODESTINO
                                        var registro = dcdList.FirstOrDefault(x => x.DCD_CONSECUTIVO == documento && x.DCD_KG_DEMOSTRADOS == (decimal)(kilpalma + kilpalmsite));

                                        // Si el registro existe, escribe el valor DCD_CODIGO en la nueva columna
                                        if (registro != null)
                                        {
                                            worksheet.Cell(row, idDcdColumn).Value = registro.DCD_CODIGO;
                                        }

                                        // Obtiene el valor de la columna ID_ENTIDAD de la fila actual
                                        string idEntidad = worksheet.Cell(row, 2).Value.ToString();

                                        // Busca el registro correspondiente en la tabla F_ENTIDAD
                                        var entidadRegistro = entidadResults.FirstOrDefault(x => x.ENTIDAD_NOMBRE == idEntidad);

                                        // Si el registro existe, escribe el valor ENT_NIT en la nueva columna
                                        if (entidadRegistro != null)
                                        {
                                            worksheet.Cell(row, nitComercializadoraColumn).Value = entidadRegistro.ENTIDAD_NIT;
                                            worksheet.Cell(row, entidadCodeColumn).Value = entidadRegistro.ENTIDAD_CODIGO;
                                        }
                                        //Busca en la tabla F_DOCCOMPROMISODESTINO el DCD_PROVEEDOR_CODIGO
                                        //string documento = worksheet.Cell(row, 6).Value.ToString();
                                        var codProvYGrupMec = dcdList.FirstOrDefault(x => x.DCD_CONSECUTIVO == documento && x.DCD_KG_DEMOSTRADOS == (decimal)(kilpalma + kilpalmsite));
                                        //Si el valor existe va a escribirlo en la tabla de excel
                                        if (codProvYGrupMec != null)
                                        {
                                            worksheet.Cell(row, codProvCodeColumn).Value = codProvYGrupMec.DCD_PROVEEDOR_CODIGO;
                                            worksheet.Cell(row, groupMerColumn).Value = codProvYGrupMec.DCD_ID_MERCADO;

                                        }

                                    }

                                    // Guarda los cambios en el archivo Excel
                                    workbook.Save();
                                    Console.WriteLine("Operación completada con éxito.");
                                    Console.WriteLine("Verifique el Excel para comprobar que los campos se actualizaron correctamente.");
                                    Console.WriteLine($"El archivo quedó en: {filePath}");
                                    Console.ReadLine();
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Se produjo un error este es el Message: {ex.Message} y  este es el StackTrace :{ex.StackTrace}");
                                Console.WriteLine($"Verificar que el archivo esté creado en esta ruta {filePath}");
                                Console.ReadLine();
                            }
                        }
                        break;

                    case 2:
                        // Lógica para insertar Finiquitos Manuales
                        Console.WriteLine("Se van a insertar los registros a la tabla de finiquitos manuales");

                        // Obtiene la ruta del archivo Excel
                        string baseDirectoryy = AppDomain.CurrentDomain.BaseDirectory;
                        string solutionDirectoryy = Path.GetFullPath(Path.Combine(baseDirectoryy, @"..\..\..\"));
                        string filePathh = Path.Combine(solutionDirectoryy, "Excel", "Manuales.xlsx");
                        Console.WriteLine("Cargando...");
                        for (int i = 0; i < 20; i++)
                        {
                            Console.Write("\u2588");
                            Thread.Sleep(500);
                        }
                        Console.WriteLine();
                        // Conexión a la base de datos Oracle
                        string connectionStringg = "DATA SOURCE=192.168.100.92:1521/SIFFANT;PERSIST SECURITY INFO=True;USER ID=fondosweb; PASSWORD=PRUEBAS2015;";

                        // Abre el archivo Excel y lee los datos de la hoja
                        using (var workbook = new XLWorkbook(filePathh))
                        {
                            var worksheet = workbook.Worksheet(1);

                            // Obtiene el rango de celdas usado en la hoja
                            var range = worksheet.RangeUsed();

                            // Inicializa la transacción en la base de datos
                            using (var connection = new OracleConnection(connectionStringg))
                            {
                                connection.Open();
                                var transaction = connection.BeginTransaction();

                                try
                                {
                                    // Itera a través de las filas del archivo Excel y las inserta en la tabla F_FINIQUITOSMANUALES
                                    for (int row = range.FirstRow().RowNumber() + 1; row <= range.LastRow().RowNumber(); row++)
                                    {
                                        // Obtén los valores de la fila actual en el archivo Excel
                                        string manComercializadora = worksheet.Cell(row, 2).Value.ToString();

                                        int manAnio;
                                        if (!int.TryParse(worksheet.Cell(row, 3).Value.ToString(), out manAnio))
                                        {
                                            manAnio = 0;
                                        }

                                        int manMes;
                                        if (!int.TryParse(worksheet.Cell(row, 4).Value.ToString(), out manMes))
                                        {
                                            manMes = 0;
                                        }

                                        string manTipoDocumento = worksheet.Cell(row, 5).Value.ToString();

                                        int manDocumento;
                                        if (!int.TryParse(worksheet.Cell(row, 6).Value.ToString(), out manDocumento))
                                        {
                                            manDocumento = 0;
                                        }

                                        double manSumaKilpalma;
                                        if (!double.TryParse(worksheet.Cell(row, 7).Value.ToString(), out manSumaKilpalma))
                                        {
                                            manSumaKilpalma = 0;
                                        }

                                        double manSumaKilpalmiste;
                                        if (!double.TryParse(worksheet.Cell(row, 8).Value.ToString(), out manSumaKilpalmiste))
                                        {
                                            manSumaKilpalmiste = 0;
                                        }

                                        string manPendienteFiniquito = worksheet.Cell(row, 9).Value.ToString();

                                        int manIdDcd;
                                        if (!int.TryParse(worksheet.Cell(row, 10).Value.ToString(), out manIdDcd))
                                        {
                                            manIdDcd = 0;
                                        }
                                        DateTime fecha;
                                        if (!DateTime.TryParse(worksheet.Cell(row, 12).Value.ToString(), out fecha))
                                        {
                                            fecha = DateTime.MinValue;
                                        }
                                        int nit;
                                        if (!int.TryParse(worksheet.Cell(row, 13).Value.ToString(), out nit))
                                        {
                                            nit = 0;
                                        }
                                        int entidadCodigo;
                                        if (!int.TryParse(worksheet.Cell(row, 14).Value.ToString(), out entidadCodigo))
                                        {
                                            entidadCodigo = 0;
                                        }
                                        int codProveedor;
                                        if (!int.TryParse(worksheet.Cell(row, 15).Value.ToString(), out codProveedor))
                                        {
                                            codProveedor = 0;
                                        }

                                        int grupoMercado;
                                        if (!int.TryParse(worksheet.Cell(row, 16).Value.ToString(), out grupoMercado))
                                        {
                                            grupoMercado = 0;
                                        }

                                        // Configura los parámetros de la consulta SQL para insertar o actualizar los datos en la tabla F_FINIQUITOSMANUALES
                                        string sqll = "MERGE INTO F_FINIQUITOSMANUALES F " +
                                                    "USING (SELECT :man_iddcd AS man_iddcd FROM DUAL) D " +
                                                    "ON (F.MAN_IDDCD = D.man_iddcd) " +
                                                    "WHEN MATCHED THEN " +
                                                    "UPDATE SET F.MAN_COMERCIALIZADORA = :man_comercializadora, " +
                                                    "F.MAN_ANIO = :man_anio, " +
                                                    "F.MAN_MES = :man_mes, " +
                                                    "F.MAN_TIPO_DOCUMENTO = :man_tipo_documento, " +
                                                    "F.MAN_DOCUMENTO = :man_documento, " +
                                                    "F.MAN_SUMA_KILPALMA = :man_suma_kilpalma, " +
                                                    "F.MAN_SUMA_KILPALMISTE = :man_suma_kilpalmiste, " +
                                                    "F.MAN_PENDIENTE_FINIQUITO = :man_pendiente_finiquito, " +
                                                    "F.MAN_APROBADAFCP = :man_aprobadafcp, " +
                                                    "F.MAN_NIT = :man_nit, " +
                                                    "F.MAN_ENTIDAD_CODIGO = :man_entidad_codigo, " +
                                                    "F.MAN_COD_PROVEEDOR = :man_cod_proveedor, " +
                                                    "F.MAN_GRUPO_MERCADO = :man_grupo_mercado " +
                                                    "WHEN NOT MATCHED THEN " +
                                                    "INSERT (MAN_ID, MAN_IDDCD, MAN_COMERCIALIZADORA, MAN_ANIO, MAN_MES, MAN_TIPO_DOCUMENTO, " +
                                                    "MAN_DOCUMENTO, MAN_SUMA_KILPALMA, MAN_SUMA_KILPALMISTE, MAN_PENDIENTE_FINIQUITO, MAN_PROCESADO, MAN_APROBADAFCP, MAN_NIT, MAN_ENTIDAD_CODIGO, MAN_COD_PROVEEDOR, MAN_GRUPO_MERCADO) " +
                                                    "VALUES (SEQ_FINIQUITOS_MANUALES.NEXTVAL, :man_iddcd, :man_comercializadora, :man_anio, " +
                                                    ":man_mes, :man_tipo_documento, :man_documento, :man_suma_kilpalma, :man_suma_kilpalmiste, " +
                                                    ":man_pendiente_finiquito, 0, :man_aprobadafcp, :man_nit, :man_entidad_codigo, :man_cod_proveedor, :man_grupo_mercado)";

                                        var parameters = new
                                        {
                                            man_iddcd = manIdDcd,
                                            man_comercializadora = manComercializadora,
                                            man_anio = manAnio,
                                            man_mes = manMes,
                                            man_tipo_documento = manTipoDocumento,
                                            man_documento = manDocumento,
                                            man_suma_kilpalma = manSumaKilpalma,
                                            man_suma_kilpalmiste = manSumaKilpalmiste,
                                            man_pendiente_finiquito = manPendienteFiniquito,
                                            man_aprobadafcp = fecha,
                                            man_nit = nit,
                                            man_entidad_codigo = entidadCodigo,
                                            man_cod_proveedor = codProveedor,
                                            man_grupo_mercado = grupoMercado
                                        };


                                        connection.Execute(sqll, parameters, transaction);
                                    }

                                    transaction.Commit();
                                    Console.WriteLine("Se insertaron los registros exitosamente");
                                    Console.ReadLine();
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("Ocurrió un error al momento de insertar los datos en la tabla");
                                }
                                finally
                                {
                                    connection.Close();
                                }
                            }
                        }
                        Console.ReadLine();
                        break;

                    case 3:
                        Console.WriteLine("Actualizando DCD_KG_DEMOSTRADOS_POLIZA en la tabla F_DOCCOMPROMISODESTINO...");
                        Console.WriteLine("Cargando...");

                        for (int i = 0; i < 20; i++)
                        {
                            Console.Write("\u2588");
                            Thread.Sleep(500);
                        }
                        Console.WriteLine();
                        // Realizar operación
                        string finiquitosSql = "SELECT * FROM F_FINIQUITOSMANUALES WHERE MAN_PROCESADO = 0";
                        string updateSql = "UPDATE F_DOCCOMPROMISODESTINO SET DCD_KG_DEMOSTRADOS_POLIZA = :suma WHERE DCD_CODIGO = :dcd_codigo AND DCD_KG_DEMOSTRADOS_POLIZA = 0";
                        string updateProcesadoSql = "UPDATE F_FINIQUITOSMANUALES SET MAN_PROCESADO = 1 WHERE MAN_ID = :man_id";
                        string connectionStringgg = "DATA SOURCE=192.168.100.92:1521/SIFFANT;PERSIST SECURITY INFO=True;USER ID=fondosweb; PASSWORD=PRUEBAS2015;";

                        using (var connection = new OracleConnection(connectionStringgg))
                        {
                            // Obtiene todos los registros de la tabla F_FINIQUITOSMANUALES con MAN_PROCESADO = 0
                            var finiquitosManuales = connection.Query<FiniquitoManualModel>(finiquitosSql).ToList();

                            // Itera sobre cada registro de F_FINIQUITOSMANUALES
                            foreach (var finiquitoManual in finiquitosManuales)
                            {
                                int manIdDcd = finiquitoManual.MAN_IDDCD;
                                decimal sumaKilpalmaKilpalmiste = finiquitoManual.MAN_SUMA_KILPALMA + finiquitoManual.MAN_SUMA_KILPALMISTE;

                                // Actualiza el campo DCD_KG_DEMOSTRADOS_POLIZA en la tabla F_DOCCOMPROMISODESTINO si cumple las condiciones
                                connection.Execute(updateSql, new { suma = sumaKilpalmaKilpalmiste, dcd_codigo = manIdDcd });

                                // Actualiza el valor de MAN_PROCESADO a 1 en la tabla F_FINIQUITOSMANUALES
                                connection.Execute(updateProcesadoSql, new { man_id = finiquitoManual.MAN_ID });
                            }
                        }
                       
                        Console.WriteLine();
                        Console.WriteLine("Actualización de DCD_KG_DEMOSTRADOS_POLIZA completada!");
                        Console.WriteLine("Presione una tecla para continuar...");
                        Console.ReadLine();
                        Console.Clear();
                        break;

                }

            } while (opcion != 0);            
        }
    }
}
