# Simulador de Investimentos em FIIs no Excel

## Sobre o projeto

Este projeto consiste em uma ferramenta de simulação de investimentos em Fundos de Investimento Imobiliário (FIIs), desenvolvida no Microsoft Excel.

O objetivo é demonstrar como fórmulas financeiras, referências estruturadas e recursos de validação de dados podem ser utilizados para transformar parâmetros de investimento em projeções de patrimônio, estimativas de dividendos e sugestões de distribuição de aportes.

A planilha foi estruturada para facilitar a interação do usuário com os parâmetros da simulação e permitir comparações entre diferentes prazos e perfis de investidor.

## Objetivos da ferramenta

O simulador busca responder às seguintes perguntas:

- Quanto investir mensalmente?
- Por quantos anos investir?
- Qual taxa de rendimento mensal utilizar na simulação?
- Qual patrimônio acumulado poderá ser projetado ao final do período?
- Qual será a estimativa de dividendos mensais com base no patrimônio projetado?

Além disso, a ferramenta apresenta cenários de acumulação para 2, 5, 10, 20 e 30 anos e distribui o aporte mensal entre diferentes categorias de FIIs, de acordo com o perfil selecionado.

## Funcionalidades

- Simulação de investimentos com aportes mensais.
- Projeção de patrimônio acumulado ao longo do tempo.
- Estimativa de dividendos mensais.
- Comparação de cenários com diferentes horizontes de investimento.
- Seleção do perfil de investidor por meio de lista suspensa.
- Distribuição percentual dos aportes entre FIIs de Papel, Tijolo e Híbridos.
- Busca automática dos percentuais por meio da função PROCV e de uma chave composta.
- Organização visual dos campos de entrada e dos resultados calculados.

## Tecnologias e recursos utilizados

- Microsoft Excel.
- Função VF (Valor Futuro).
- Função PROCV.
- Validação de dados.
- Intervalos nomeados.
- Referências absolutas e relativas.
- Chaves compostas para consulta de dados.
- Formatação de células e organização de interface.

## Como os cálculos funcionam

### 1. Função VF — Valor Futuro

A função VF é utilizada para projetar o patrimônio acumulado a partir dos aportes periódicos e da taxa de rendimento mensal informada.

Seus argumentos permitem definir:

- **Taxa:** rendimento por período, neste caso mensal.
- **Nper:** número total de períodos, calculado a partir do prazo em anos multiplicado por 12.
- **Pgto:** valor do aporte realizado mensalmente.
- **Vp:** valor presente ou saldo inicial, quando aplicável.
- **Tipo:** indica se os pagamentos são considerados no início ou no final de cada período.

O sinal dos aportes deve ser compatível com a convenção financeira da função, para que o patrimônio seja apresentado com o sinal esperado.

A mesma lógica é utilizada para projetar os diferentes cenários de investimento.

### 2. Estimativa de dividendos

A ferramenta utiliza o patrimônio projetado e a taxa de dividendos mensais considerada na simulação para estimar os recebimentos mensais.

Essa estimativa depende dos parâmetros informados e não representa uma garantia de rentabilidade ou de distribuição de dividendos.

### 3. Função PROCV e chave composta

Para identificar o percentual de alocação correspondente a cada tipo de fundo, a planilha combina o perfil do investidor e a categoria do FII em uma única chave.

Por exemplo:

`Moderado - PAPEL`

A chave é comparada com a primeira coluna da tabela auxiliar. Quando encontra uma correspondência exata, a função PROCV retorna o percentual associado à combinação encontrada.

Exemplo da fórmula utilizada na planilha:

```excel
=PROCV($B$42&" - "&A46;Planilha1!$A$2:$D$10;4;FALSO)
```

Nessa fórmula:

- `$B$42` representa o perfil selecionado.
- `A46` representa a categoria do fundo.
- `&" - "&` concatena as informações para formar a chave composta.
- `Planilha1!$A$2:$D$10` corresponde à tabela auxiliar de consulta.
- `4` indica que o resultado será retornado da quarta coluna da tabela.
- `FALSO` exige uma correspondência exata.

Ao alterar o perfil ou copiar a fórmula para as demais categorias, a planilha consegue buscar os percentuais correspondentes automaticamente.

## Intervalos nomeados

Os intervalos nomeados facilitam a leitura das fórmulas e tornam a manutenção da planilha mais simples, reduzindo a dependência de referências diretas a endereços de células.

| Intervalo nomeado | Finalidade |
|---|---|
| `aporte` | Identifica o valor do investimento mensal utilizado na simulação. |
| `taxa_mensal` | Identifica a taxa de rendimento mensal utilizada nos cálculos. |

Os nomes acima devem corresponder aos intervalos efetivamente definidos no Gerenciador de Nomes do Excel.

## Distribuição por perfil de investidor

Os percentuais configurados na tabela auxiliar são:

| Perfil | Papel | Tijolo | Híbrido | Total |
|---|---:|---:|---:|---:|
| Conservador | 30% | 50% | 20% | 100% |
| Moderado | 32% | 35% | 33% | 100% |
| Agressivo | 50% | 35% | 15% | 100% |

Esses percentuais são parâmetros ilustrativos utilizados no projeto para demonstrar o funcionamento da distribuição dos aportes. Não constituem recomendação de investimento.

## Como utilizar a planilha

1. Abra o arquivo no Microsoft Excel.
2. Informe os parâmetros necessários à simulação, como o aporte mensal, o prazo e as taxas utilizadas.
3. Consulte o patrimônio acumulado e a estimativa de dividendos mensais.
4. Compare os cenários projetados para diferentes horizontes de investimento.
5. Selecione outro perfil de investidor na lista suspensa.
6. Verifique a atualização automática dos percentuais de alocação para cada categoria de FII.

## Desenvolvimento e aprendizados

O desenvolvimento deste projeto permitiu praticar a aplicação de fórmulas financeiras no Excel, organizar dados para consultas automatizadas e estruturar uma ferramenta de simulação orientada a uma necessidade de negócio.

Entre os principais aprendizados estão a utilização da função VF para projeções financeiras, o emprego do PROCV com chaves compostas, o uso de intervalos nomeados e a construção de uma interface que facilita a interpretação dos resultados.

A estrutura do projeto foi desenvolvida a partir da proposta do desafio, com uma organização visual própria para os parâmetros, resultados e tabela de distribuição por perfil.

## Premissas e limitações

Os resultados apresentados são estimativas matemáticas baseadas nos parâmetros informados pelo usuário. A simulação não garante rentabilidade futura nem considera automaticamente impostos, taxas, inflação, oscilações de mercado ou mudanças nas distribuições dos fundos, a menos que esses fatores sejam incorporados aos cálculos.

O projeto possui finalidade educacional e demonstra recursos do Excel aplicados à simulação financeira.

---

**Projeto desenvolvido para fins de estudo e portfólio.**
