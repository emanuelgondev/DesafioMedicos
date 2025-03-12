using Consultas.Models;
using MathNet.Numerics;
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
        //ExercicioQuatro();
        ExercicioCinco();
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
                double valorConsulta = (double)linha.GetCell(12).NumericCellValue;

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

        var nomeDosPacientes = consultas.Select(n => n.NomePaciente).ToList();
        var pacientesDoPeriodo = consultas.Where(c => c.DataConsulta >= dataInicio && c.DataConsulta <= dataFinal)
            .DistinctBy(n => n.NomePaciente).ToList();
        foreach (var nomes in nomeDosPacientes)
        {
            Console.WriteLine($"Paciente: {nomes}");
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
        var nomeDosMedicoas = consultas.Select(n => n.NomeMedico).ToList();
        var medicosTrabalhando = consultas.DistinctBy(n => n.NomeMedico).ToList();
        foreach (var nomes in nomeDosMedicoas)
        {
            Console.WriteLine($"Médico: {nomes}");
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
            Console.WriteLine($"Médico: {nomes.NomeMedico} - {string.Join(", ", nomes.Especialidades)}");
        }

    }

    //4 – Liste o total em valor de consulta que receberemos. Some o valor de todas as consultas.
    // Total em dinheiro das consultas do dia 27/03 – 31/03 – R$ 10.000.
    // Cardiologia - R$ 3.000
    // Pediatra – R$ 2.000
    // Oftalmologia – R$ 5.000

    static void ExercicioQuatro()
    {
        // 4 – Liste o total em valor de consulta que receberemos. Some o valor de todas as consultas.
        // Depois liste o valor por especialidade.

        // Exemplo

        // Total em dinheiro das consultas do dia 27/03 – 31/03 – R$ 10.000.
        // Cardiologia - R$ 3.000
        // Pediatra – R$ 2.000
        // Oftalmologia – R$ 5.000

        var nomeEspecialidades = consultas.GroupBy(e => e.Especialidade).Select(a => new
        {
            especialidade = a.Key,
            valor = a.Sum(b => b.ValorConsulta)
        });
        Console.WriteLine($"Total em consultas do dia 27/03 – 31/03: R$ {nomeEspecialidades.Sum(valor => valor.valor):c}");
        foreach (var especialidades in nomeEspecialidades)
        {
            Console.WriteLine($"{especialidades.especialidade} - R$ {especialidades.valor:c}");

        }
    }
    static void ExercicioCinco()
    {
        // 5 – Para o dia 30/03. Quantas consultas vão ser realizadas? Quantas são Particular? Liste para esse dia os horários de consulta de cada médico e suas especialidades.

        // Exemplo

        // Para o dia 30/03 – Total de 10 consultas. 7 particular e 3 convênios.

        // Wagner – Cardiologista : 08:00, 09:00, 16:00
        // Wagner – Oftalmologia: 12:00
        // Tatiana – Pediatra : 09:00, 10:00, 13:00
        DateTime data = new DateTime(2023, 03, 30);

        var consultasNaData = consultas.Where(c => c.DataConsulta == data);

        var particular = consultasNaData.Where(c => c.Particular == true);

        Console.WriteLine($"Total de consultas para o dia 30/03: {consultasNaData.Count()}");
        Console.WriteLine($"Dessas, {particular.Count()} são particulares");

        var saidaFinal = consultasNaData.GroupBy(c => c.NomeMedico)
        .Select(c => new
        {
            nome = c.Key,
            especialidades = c.Select(c => c.Especialidade).Distinct(),
            horaConsulta = c.Select(c => c.HoraConsulta).Distinct()
        });

        foreach (var medico in saidaFinal)
        {
            Console.WriteLine($"{medico.nome} - {string.Join(", ", medico.especialidades)}: terá uma consulta as {string.Join(", ", medico.horaConsulta)}");
        }
    }
}
