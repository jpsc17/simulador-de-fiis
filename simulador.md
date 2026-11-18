# 🏢 Simulador de Investimentos em Fundos Imobiliários (FIIs)

Ferramenta construída no Microsoft Excel com visual de aplicativo para simulação de acúmulo de patrimônio, projeção de renda passiva mensal (dividendos) e alocação estratégica de aportes por perfil de investidor em seis segmentos de Fundos Imobiliários.

---

## 🎯 1. As Cinco Perguntas de Negócio (e Onde Encontrá-las)

Na aba principal **`Simulador_FIIs`**, o bloco **1. As 5 Perguntas de Negócio** (`B6:D11`) responde diretamente às questões fundamentais do investidor:

1. **Quanto investir por mês?** Respondida na célula editável `D7` (Intervalo nomeado: `aporte`). Valor padrão de exemplo: `R$ 1.000,00`.
2. **Por quantos anos?** Respondida na célula editável `D8` (Intervalo nomeado: `prazo_anos`). Valor padrão de exemplo: `10 anos`.
3. **Qual a taxa de rendimento mensal?** Respondida na célula calculada `D9` (Intervalo nomeado: `taxa_mensal`), espelhada da configuração de rendimento médio da carteira (`D16`, configurada em `0,85% a.m.`).
4. **Quanto de patrimônio vai acumular?** Calculada automaticamente em cinza na célula `D10` (`patrimonio_acumulado`) via função `VF`.
5. **Quanto vai receber de dividendos por mês?** Calculada automaticamente em cinza na célula `D11` (`dividendos_mensais`), multiplicando o patrimônio acumulado pela taxa mensal.

---

## 🧮 2. Como as Funções `VF` e `PROCV` Entram nos Cálculos

### Função `VF` (Valor Futuro)
Utilizada para projetar os juros compostos dos aportes mensais recorrentes ao longo do tempo:
```excel
=VF(taxa_mensal; prazo_anos*12; -aporte; 0)