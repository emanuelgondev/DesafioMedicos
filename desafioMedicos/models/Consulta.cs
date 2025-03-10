using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace Consultas.Models
{
    public class Consulta
    {
        public DateTime? DataConsulta { get; protected set; }
        public string HoraConsulta { get; protected set; }
        public string NomePaciente { get; protected set; }
        public string NumeroTelefone { get; protected set; }
        public long Cpf { get; protected set; }
        public string Rua { get; protected set; }
        public string Cidade { get; protected set; }
        public string Estado { get; protected set; }
        public string Especialidade { get; protected set; }
        public string NomeMedico { get; protected set; }
        public bool Particular { get; protected set; }
        public long NumeroCarteirinha { get; protected set; }
        public double ValorConsulta { get; protected set; }

        public Consulta(
            DateTime dataConsulta,
            string horaConsulta,
            string nomePaciente,
            string numeroTelefone,
            long cpf,
            string rua,
            string cidade,
            string estado,
            string especialidade,
            string nomeMedico,
            bool particular,
            long numeroCarteirinha,
            double valorConsulta
            )
        {
            SetDataConsulta(dataConsulta);
            SetHoraConsulta(horaConsulta);
            SetNomePaciente(nomePaciente);
            SetNumeroTelefone(numeroTelefone);
            SetCpf(cpf);
            SetRua(rua);
            SetCidade(cidade);
            SetEstado(estado);
            SetEspecialidade(especialidade);
            SetNomeMedico(nomeMedico);
            SetParticular(particular);
            SetNumeroCarteirinha(numeroCarteirinha);
        }

        public void SetDataConsulta(DateTime dataConsulta)
        {
            if (dataConsulta < DateTime.MinValue)
                throw new ArgumentException("Data da consulta não pode ser anterior a data de hoje.");
            DataConsulta = dataConsulta;
        }

        public void SetHoraConsulta(string horaConsulta)
        {
            if (string.IsNullOrWhiteSpace(horaConsulta) || !TimeSpan.TryParse(horaConsulta, out TimeSpan parsedTime))
                throw new ArgumentException("Hora da consulta inválida.");
            HoraConsulta = horaConsulta;
        }

        public void SetNomePaciente(string nomePaciente)
        {
            if (string.IsNullOrWhiteSpace(nomePaciente))
                throw new ArgumentException("Nome do paciente não pode estar vazio.");
            NomePaciente = nomePaciente;
        }

        public void SetNumeroTelefone(string numeroTelefone)
        {
            // if (string.IsNullOrWhiteSpace(numeroTelefone))
            //     throw new ArgumentException("Número de telefone inválido.");
            NumeroTelefone = numeroTelefone;
        }

        public void SetCpf(long cpf)
        {
            if (cpf <= 0)
                throw new ArgumentException("CPF inválido.");
            Cpf = cpf;
        }

        public void SetRua(string rua)
        {
            if (string.IsNullOrWhiteSpace(rua))
                throw new ArgumentException("Rua do paciente não pode estar vazia.");
            Rua = rua;
        }

        public void SetCidade(string cidade)
        {
            if (string.IsNullOrWhiteSpace(cidade))
                throw new ArgumentException("Cidade do paciente não pode estar vazia.");
            Cidade = cidade;
        }

        public void SetEstado(string estado)
        {
            if (string.IsNullOrWhiteSpace(estado))
                throw new ArgumentException("Estado do paciente inválido.");
            Estado = estado;
        }

        public void SetEspecialidade(string especialidade)
        {
            if (string.IsNullOrWhiteSpace(especialidade))
                throw new ArgumentException("Especialidade do médico não pode estar vazia.");
            Especialidade = especialidade;
        }

        public void SetNomeMedico(string nomeMedico)
        {
            if (string.IsNullOrWhiteSpace(nomeMedico))
                throw new ArgumentException("Nome do médico não pode estar vazio.");
            NomeMedico = nomeMedico;
        }

        public void SetParticular(bool particular)
        {
            Particular = particular;
        }

        public void SetNumeroCarteirinha(long numeroCarteirinha)
        {
            if (numeroCarteirinha <= 0)
                throw new ArgumentException("Número da carteirinha inválido.");
            NumeroCarteirinha = numeroCarteirinha;
        }

        public void SetValorConsulta(double valorConsulta)
        {
            ValorConsulta = valorConsulta;
        }
    }
}