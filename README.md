# 🏠 Ela Investe — Simulador de Fundos Imobiliários (FIIs) em Excel

Planilha em Excel para simular investimentos em **Fundos Imobiliários**: quanto investir por mês, por quanto tempo, qual o patrimônio acumulado e quanto isso rende em **dividendos mensais**. Também ajuda a planejar metas de vida e a entender o impacto de uma pausa nos aportes.

Projeto desenvolvido como desafio do bootcamp da **[DIO](https://www.dio.me/)**, com base na aula e com personalizações minhas.

> ⚠️ Material educativo. Não é recomendação de investimento. Os números são simulações e rendimentos reais variam.

---

## 🎯 Objetivo

Aplicar conceitos de Excel na construção de uma ferramenta prática de simulação, automatizando cálculos como valor total investido, patrimônio acumulado e dividendos mensais, para apoiar decisões mais informadas.

## 📚 O que aprendi / pratiquei

- Criar uma ferramenta de simulação de investimentos no Excel
- Cálculos financeiros: juros compostos, rendimento mensal e dividendos
- Funções financeiras (`FV`, `PMT`) e de busca (`VLOOKUP`)
- **Nomes definidos** para deixar as fórmulas legíveis
- **Validação de dados** (lista suspensa de perfil)
- Gráficos e organização visual de uma planilha
- Documentar um projeto técnico e publicá-lo no GitHub

---

## 🗂️ Estrutura da planilha

O arquivo `.xlsx` deste repositório tem 5 abas:

| Aba | Para que serve |
|---|---|
| **Planilha2** | Simulador principal: configurações, investimento mensal, cenários de 2 a 30 anos e sugestão de FIIs por perfil |
| **Sobre Investimentos** | Conteúdo educativo: o que é investir, tipos de investimento, independência financeira e foco em FIIs |
| **Metas de Vida** | Quanto guardar por mês para cada objetivo (viagem, casa, negócio, aposentadoria) |
| **Pausa na Carreira** | Compara o patrimônio com e sem pausa nos aportes e calcula quanto guardar para recuperar o tempo |
| **Planilha-apoio** | Tabela de percentuais por perfil e tipo de FII, usada nas buscas |

### 1. Simulador principal (Planilha2)

**Entradas:** salário, rendimento da carteira (dividend yield mensal), aporte mensal, anos de investimento e taxa de rendimento mensal.

**Saídas:** patrimônio acumulado e dividendos mensais estimados.

Exemplo com os valores da planilha:

| Entrada | Valor |
|---|---|
| Aporte mensal | R$ 1.200 |
| Prazo | 10 anos |
| Taxa mensal | 1,079% |
| Rendimento da carteira (dividendos) | 0,6% ao mês |

| Resultado | Valor |
|---|---|
| Patrimônio acumulado | ≈ R$ 291.941 |
| Dividendos mensais | ≈ R$ 1.751 |

**Tabela de cenários** (mesmo aporte, prazos diferentes):

| Prazo | Patrimônio | Dividendo mensal |
|---|---|---|
| 2 anos | R$ 32.673 | R$ 196 |
| 5 anos | R$ 100.532 | R$ 603 |
| 10 anos | R$ 291.941 | R$ 1.751 |
| 20 anos | R$ 1.350.238 | R$ 8.101 |
| 30 anos | R$ 5.186.604 | R$ 31.120 |

Também há uma **sugestão de investimento de 30% do salário**.

### 2. Sugestão de carteira por perfil

Ao escolher **Conservador, Moderado ou Agressivo** em uma lista suspensa, a planilha distribui o aporte mensal entre os tipos de FII: Papel, Tijolo, Híbridos, FoFs, Desenvolvimento e Hotelarias. A distribuição vem da aba `Planilha-apoio`, via `VLOOKUP` com chave `PERFIL-TIPO`.

### 3. Metas de Vida

Informe o valor da meta e o prazo, e a planilha calcula quanto guardar por mês, quanto sai do seu bolso, quanto vem de rendimento e quanto isso pesa no salário. No final, avisa se a soma das metas cabe na sugestão de investimento.

### 4. Pausa na Carreira

Simula três cenários: (A) sem pausa, (B) pausa e volta com o mesmo aporte, (C) pausa e volta com aporte menor. Mostra a diferença no patrimônio final e quanto guardar por mês ao retomar para compensar.

---

## 🧮 Fórmulas principais

| Cálculo | Fórmula |
|---|---|
| Patrimônio acumulado | `=FV(taxa_mensal; qtd_anos*12; aporte*-1)` |
| Dividendos mensais | `=patrimonio * Rendimento_carteira` |
| Sugestão de investimento | `=salario * 30%` |
| Percentual do perfil | `=VLOOKUP(perfil&"-"&tipo; 'Planilha-apoio'!A:D; 4; )` |
| Aporte necessário para uma meta | `=-PMT(taxa; meses; 0; valor_meta)` |

**Nomes definidos:** `salario`, `aporte`, `qtd_anos`, `taxa_mensal`, `patrimonio`, `Rendimento_carteira`, `sugestao_investimento`.

**Premissas:** aportes feitos no fim de cada mês, taxa constante e sem considerar impostos, inflação ou variação do preço das cotas.


Projeto baseado no laboratório da [DIO](https://www.dio.me/), com adaptações e funcionalidades extras feitas por mim.
