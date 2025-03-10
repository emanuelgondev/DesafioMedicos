using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace medicos.models
{
    public class Medico
    {
        public string NomeMedico { get; protected set; }
        public List<Medico> Especialidades { get; protected set; }

        public Medico(
            string nomeMedico,
            List<Medico> especialidades
        )
        {
            SetMedico(nomeMedico);
        }

        protected void SetMedico(string nomeMedico)
        {
            if (string.IsNullOrEmpty(nomeMedico))
                throw new ArgumentException("Nome do médico não pode ser vazio.");

            NomeMedico = nomeMedico;
        }
    }
}