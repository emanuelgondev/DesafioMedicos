using Consultas.Models;
using NPOI.SS.UserModel;
using System.Text.RegularExpressions;
internal class Program
{
    private static string caminhoArquivo = Path.Combine(Environment.CurrentDirectory, "DesafioMedicos.xlsx");
    private static List<Consulta> consultas = [];

    private static void Main(string[] args)
    {
        Console.Clear();

        ImportarDadosPlanilha();
        //ExercicioUm();
        //ExercicioDois();
        //ExercicioTres();
        ExercicioQuatro();
    }
    private static void ImportarDadosPlanilha()
    {
        try
        {
            IWorkbook pastaDesafio = WorkbookFactory.Create(caminhoArquivo);

            ISheet planilha = pastaDesafio.GetSheetAt(0);
            for (int i = 1; i < planilha.PhysicalNumberOfRows; i++)
            {
                IRow linha = planilha.GetRow(i);

                DateTime dataConsulta = DateTime.Parse(linha.GetCell(0).StringCellValue);
                string horaConsulta = linha.GetCell(1).StringCellValue;
                string nomePaciente = linha.GetCell(2).StringCellValue;
                string numeroTelefone = linha.GetCell(3)?.StringCellValue;
                // cpf = (long)linha.GetCell(4).NumericCellValue;
                long cpf = Convert.ToInt64(Regex.Replace(linha.GetCell(4).StringCellValue, @"\D", ""));
                string rua = linha.GetCell(5).StringCellValue;
                string cidade = linha.GetCell(6).StringCellValue;
                string estado = linha.GetCell(7).StringCellValue;
                string especialidade = linha.GetCell(8).StringCellValue;
                string nomeMedico = linha.GetCell(9).StringCellValue;
                bool particular = linha.GetCell(10).StringCellValue == "Sim" ? true : false;
                long numeroCarteirinha = (long)linha.GetCell(11).NumericCellValue;
                double valorConsulta = linha.GetCell(12).NumericCellValue;

                Consulta consulta = new(dataConsulta, horaConsulta, nomePaciente, numeroTelefone, cpf, rua, cidade, estado, especialidade, nomeMedico, particular, numeroCarteirinha, valorConsulta);

                consultas.Add(consulta);
            }

        }
        catch (Exception e)
        {
            Console.WriteLine(e.Message);
        }
    }

    static void ExercicioUm()
    {
        //  1 – Liste ao total quantos pacientes temos para atender do dia 27/03 até dia 31/03. Sem repetições.

        // Exemplo
        // Aline
        // Aline
        // João
        // Arthur

        // Total: 3 pacientes

        DateTime dataInicio = new DateTime(2023, 03, 27);
        DateTime dataFinal = new DateTime(2023, 03, 31);

        var pacientesDoPeriodo = consultas.Where(c => c.DataConsulta >= dataInicio && c.DataConsulta <= dataFinal)
            .ToList();
        var nomeDosPacientes = pacientesDoPeriodo.GroupBy(n => n.NomePaciente).Select(p => p.First()).ToList();
        foreach (var nomes in nomeDosPacientes)
        {
            Console.WriteLine($"Paciente: {nomes.NomePaciente}");
        }
        Console.WriteLine($"Total: {pacientesDoPeriodo.Count()} pacientes");
    }
    static void ExercicioDois()
    {
        // 2 – Liste ao total quantos médicos temos trabalhando em nosso consultório. Conte a quantidade de médicos sem repetições. 

        // Exemplo
        // Aline
        // Aline
        // João
        // Arthur

        // Total: 3 médicos
        var medicosTrabalhando = consultas.GroupBy(n => n.NomeMedico).Select(m => m.First()).ToList();
        foreach (var nomes in medicosTrabalhando)
        {
            Console.WriteLine($"Médico: {nomes.NomeMedico}");
        }
        Console.WriteLine($"Total: {medicosTrabalhando.Count()} medicos trabalhando.");

    }
    static void ExercicioTres()
    {
        // 3 – Liste o nome dos médicos e suas especialidades.

        // Exemplo
        // Wagner – Cardiologia,  Oftalmologia.
        var medicos = consultas.GroupBy(n => n.NomeMedico).Select(grupo => new
        {
            NomeMedico = grupo.Key,
            Especialidades = grupo.Select(c => c.Especialidade).Distinct().ToList()
        })
        .ToList();
        foreach (var nomes in medicos)
        {
            Console.WriteLine($"Médico: {nomes.NomeMedico} - {string.Join(", ",nomes.Especialidades)}");
        }

    }
    static void ExercicioQuatro(){ //ValorConsulta criado sem Exception

    }
}
