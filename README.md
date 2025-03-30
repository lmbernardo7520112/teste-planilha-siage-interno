# SIAGE INTERNO ECI LUIS RAMALHO - Sistema Integrado de Análise e Gestão Escolar

![SIAGE Logo](app/core/static/images/siage_interno.png)

O SIAGE é um sistema avançado para geração automatizada de planilhas de notas e relatórios escolares, desenvolvido para otimizar o trabalho de instituições educacionais. Este projeto demonstra habilidades avançadas em Python, manipulação de planilhas Excel com OpenPyXL, e criação de dashboards analíticos.

## ✨ Recursos Principais

- **Geração automatizada** de planilhas de notas completas
- **Dashboards interativos** com visualização de dados educacionais
- **Análise de desempenho** por turma, disciplina e aluno
- **Controle de situação acadêmica** (ativos, transferidos, desistentes)
- **Cálculos automáticos** de médias, aprovações e indicadores educacionais

## 🛠️ Tecnologias Utilizadas

- **Python 3.10+**
- **OpenPyXL** - Para manipulação avançada de planilhas Excel
- **Logging** - Para registro de atividades do sistema
- **Pathlib** - Para manipulação segura de caminhos de arquivos
- **JSON** - Para armazenamento e leitura de dados estruturados

## 📊 Estrutura do Projeto

O sistema gera uma planilha Excel complexa com múltiplas abas contendo:

1. **Abas de Disciplinas**: Uma para cada disciplina com:
   - Notas bimestrais
   - Cálculo de médias
   - Situação do aluno
   - Dashboard de desempenho da turma

2. **Aba SEC**: Contendo:
   - Status dos alunos (Ativo/Transferido/Desistente)
   - Dashboards de análise de evasão
   - Taxas de aprovação por turma

3. **Aba Boletim Consolidado**: Resumo completo de todas as disciplinas

4. **Abas Adicionais**: Para relatórios individuais e frequência

## 🎨 Recursos de Design

- **Formatação profissional** com cores e estilos padronizados
- **Logotipo institucional** em todas as abas
- **Gráficos automáticos** para visualização de dados
- **Layout responsivo** que se adapta ao número de alunos

## ⚙️ Configuração

O sistema é altamente configurável através do arquivo `config.py` que permite:

- Definir as disciplinas oferecidas
- Personalizar cores e estilos
- Ajustar fórmulas de cálculo
- Modificar estruturas de relatórios

## 📈 Indicadores Calculados

O sistema calcula automaticamente:

- Taxas de aprovação/reprovação por turma e disciplina
- Médias bimestrais e finais
- Percentual de alunos com desempenho acima da média
- Índices de evasão escolar
- Situação acadêmica de cada aluno

## 🚀 Como Executar

1. Clone o repositório
2. Instale as dependências: `pip install openpyxl`
3. Configure os arquivos JSON com os dados dos alunos
4. Execute o script principal

