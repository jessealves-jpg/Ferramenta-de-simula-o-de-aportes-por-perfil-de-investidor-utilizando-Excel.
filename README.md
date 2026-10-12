# Ferramenta de Simulação de Aportes por Perfil de Investidor

## Sobre o projeto

Este projeto foi desenvolvido em Microsoft Excel durante o acompanhamento de um curso, com o objetivo de simular aportes financeiros, estimar o crescimento do patrimônio ao longo do tempo e calcular uma sugestão de distribuição dos investimentos de acordo com o perfil do investidor.

A ferramenta utiliza fórmulas financeiras, consultas de dados e uma tabela de apoio para apresentar simulações e distribuir o aporte mensal entre diferentes categorias de Fundos de Investimento Imobiliário (FIIs).

## Objetivos

- Praticar o uso de fórmulas financeiras no Excel.
- Consultar informações em uma tabela de apoio por meio da função PROCV.
- Calcular estimativas de patrimônio futuro com a função VF.
- Utilizar intervalos nomeados para facilitar a leitura das fórmulas.
- Simular a distribuição de aportes conforme o perfil selecionado.
- Apresentar cenários financeiros para diferentes períodos.

## Perguntas que a ferramenta responde

**1. Quanto posso investir por mês?**

A planilha apresenta uma sugestão de investimento equivalente a 30% do salário informado, permitindo comparar esse valor com o aporte mensal escolhido.

**2. Quanto posso acumular ao longo do tempo?**

Calcula uma estimativa do patrimônio futuro com base no aporte mensal, no prazo e na taxa de rendimento informada.

**3. Quanto posso receber mensalmente em dividendos?**

Estima os dividendos mensais multiplicando o patrimônio acumulado pelo rendimento mensal da carteira.

**4. Como meu patrimônio pode evoluir em diferentes períodos?**

Apresenta cenários de acumulação para 2, 5, 10, 20 e 30 anos, permitindo comparar os resultados estimados.

**5. Como distribuir meu aporte mensal entre as categorias de FIIs?**

Consulta os percentuais associados ao perfil selecionado e calcula o valor destinado a cada categoria de investimento.

## Fórmulas utilizadas

### 1. PROCV — Consulta dos percentuais por perfil

Fórmula utilizada:

```excel
=PROCV($D$29&"-"&C33;TB_APOIO!$A:$D;4;FALSO)
```

A função PROCV consulta a tabela de apoio `TB_APOIO` para localizar a combinação do perfil selecionado na célula `D29` com a categoria de FII indicada em `C33`.

O operador `&` concatena as duas informações, utilizando um hífen como separador. A função procura a combinação na primeira coluna da tabela de apoio e retorna o valor correspondente da quarta coluna.

O argumento `FALSO` determina que a pesquisa seja realizada por correspondência exata.

Dessa forma, os percentuais retornados dependem do perfil selecionado e dos dados cadastrados na tabela de apoio.

### 2. VF — Cálculo do valor futuro

Fórmula utilizada:

```excel
=VF($E$17;$A22*12;$E$15*-1)
```

A função VF calcula o valor futuro estimado dos aportes periódicos, considerando os seguintes parâmetros:

- `E17`: taxa de rendimento mensal.
- `A22*12`: prazo convertido de anos para meses.
- `E15*-1`: aporte mensal convertido em valor negativo para representar a saída de dinheiro no cálculo financeiro.

O resultado representa uma estimativa do patrimônio acumulado ao final do período calculado, considerando a taxa de rendimento e os aportes informados.

### 3. Cálculo dos dividendos mensais

Fórmula utilizada:

```excel
=patrimonio*rendimento_carteira
```

Essa fórmula multiplica o patrimônio acumulado pelo rendimento mensal da carteira, estimando o valor dos dividendos mensais.

O resultado depende do patrimônio calculado e da taxa de rendimento definida na planilha.

## Intervalos nomeados

A planilha utiliza os seguintes nomes nas fórmulas:

| Nome | Finalidade |
|---|---|
| `TB_APOIO` | Referência à tabela de apoio consultada pela função PROCV. |
| `patrimonio` | Referência ao valor do patrimônio acumulado utilizado no cálculo dos dividendos. |
| `rendimento_carteira` | Referência ao rendimento mensal utilizado para estimar os dividendos. |

Esses nomes ajudam a identificar as referências utilizadas nos cálculos e facilitam a compreensão das fórmulas.

## Origem dos percentuais

Os percentuais de distribuição dos investimentos foram copiados da ferramenta original do Expert, utilizada como referência durante o curso.

A tabela de apoio reúne os percentuais por perfil e categoria de FII. A função PROCV consulta esses dados para retornar o percentual correspondente à combinação selecionada.

A distribuição calculada para cada perfil deve totalizar 100% do aporte mensal.

## Exemplo de simulação

Os valores a seguir são fictícios e servem apenas para demonstrar o funcionamento da ferramenta.

| Parâmetro | Valor de exemplo |
|---|---:|
| Salário mensal | R$ 4.200,00 |
| Sugestão de investimento (30%) | R$ 1.260,00 |
| Aporte mensal simulado | R$ 200,00 |
| Prazo de investimento | 5 anos |
| Taxa de rendimento mensal | 1,08% |
| Perfil selecionado | Agressivo |
| Patrimônio acumulado estimado | R$ 16.755,38 |

### Distribuição do aporte — Perfil Agressivo

| Categoria de FII | Percentual sugerido | Valor mensal |
|---|---:|---:|
| Papel | 50% | R$ 100,00 |
| Tijolo | 10% | R$ 20,00 |
| Híbridos | 5% | R$ 10,00 |
| FOFs | 5% | R$ 10,00 |
| Desenvolvimento | 20% | R$ 40,00 |
| Hotelarias | 10% | R$ 20,00 |
| **Total** | **100%** | **R$ 200,00** |

## Evidências de funcionamento

Os prints abaixo demonstram o funcionamento da planilha em diferentes perfis de investidor. A comparação permite visualizar a alteração dos percentuais e dos valores destinados a cada categoria.

### Perfil Agressivo

Adicionar aqui a captura de tela da planilha preenchida com o perfil Agressivo.

### Segundo perfil de investidor

Adicionar aqui a captura de tela da mesma simulação após selecionar outro perfil.

## Diferenças em relação à ferramenta do Expert

O projeto foi desenvolvido acompanhando as etapas e os recursos apresentados pelo instrutor durante o curso, reproduzindo a estrutura e os cálculos ensinados.

As personalizações realizadas foram principalmente visuais, incluindo alterações nas cores, nas linhas e na formatação da planilha, com o objetivo de melhorar a apresentação e a leitura das informações.

Os percentuais utilizados foram mantidos conforme a ferramenta original do Expert.

## Tecnologias e recursos utilizados

- Microsoft Excel
- Função VF para cálculos financeiros
- Função PROCV para consultas de dados
- Intervalos nomeados
- Tabela de apoio para distribuição percentual
- Simulação de cenários financeiros

## Observações

Os valores apresentados são fictícios e têm finalidade demonstrativa. As simulações dependem das taxas e dos aportes informados e não garantem rentabilidade futura.

Os percentuais apresentados reproduzem os parâmetros da ferramenta de referência e não constituem recomendação individual de investimento.

## Autor

Jesse Alves de Oliveira

[GitHub](https://github.com/jessealves-jpg)
