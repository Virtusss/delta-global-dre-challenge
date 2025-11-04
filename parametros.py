# Arquivo de Parâmetros para geração da DRE

# Modifique os valores abaixo conforme necessário

# Taxa de imposto sobre o lucro (em porcentagem)
# Exemplo: 10 para 10%, 15 para 15%, etc.
taxa_imposto = 30

# Detecção automática do período da DRE
# Se True: encontra automaticamente as datas mínima e máxima nas planilhas
# Se False: usa os parâmetros periodo_inicio e periodo_final definidos abaixo
auto_detectar_periodo = True
# Período específico (usado apenas se auto_detectar_periodo = False)
# Formato: MM/YY (ex: '01/24' para janeiro de 2024)
periodo_inicio = '03/24'  # Mês/Ano inicial
periodo_final = '01/24'   # Mês/Ano final

# Vida útil dos ativos por tipo (em anos)
# Define o número de anos para depreciar cada tipo de ativo
# A chave deve corresponder EXATAMENTE ao texto na coluna "Descrição" da aba Investimentos
vida_util_ativos = {
    'Expansão': 3,
    'Equipamento': 3,
    'Software': 5,
}
# Valor padrão para vida útil quando o tipo não é encontrado
vida_util_padrao = 5