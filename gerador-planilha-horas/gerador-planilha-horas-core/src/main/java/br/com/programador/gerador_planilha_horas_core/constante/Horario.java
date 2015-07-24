package br.com.programador.gerador_planilha_horas_core.constante;

public enum Horario {
	ENTRADA("08:00:00"), SAIDA_ALMOCO("12:00:00"), RETORNO_ALMOCO("13:00:00"), SAIDA("17:00:00");

	private Horario(String horario) {
		this.horario = horario;
	}

	private String horario;

	public String getHorario() {
		return horario;
	}

	public void setHorario(String horario) {
		this.horario = horario;
	}

}
