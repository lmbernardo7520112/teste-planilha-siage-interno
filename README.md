# SIAGE INTERNO ECI LUIS RAMALHO - Sistema Integrado de Análise e Gestão Escolar

[![License][License-shield]][License-url]
[![Contributors][Contributors-shield]][Contributors-url]
[![Forks][Forks-shield]][Forks-url]
[![Stargazers][Stars-shield]][Stars-url]
[![Issues][Issues-shield]][Issues-url]

[![Python][Python-shield]][Python-url]
[![OpenPyXL][OpenPyXL-shield]][OpenPyXL-url]

Sistema avançado para geração automatizada de planilhas de notas e relatórios escolares detalhados, projetado para otimizar processos e análises em instituições educacionais como a ECI Luis Ramalho.

## ✨ Sobre o Projeto

O SIAGE INTERNO foi desenvolvido para simplificar e automatizar a complexa tarefa de compilar notas, calcular médias, analisar o desempenho dos alunos e gerar relatórios consolidados. Utilizando Python e a biblioteca OpenPyXL, o sistema processa dados de entrada (presumivelmente em JSON) e produz uma planilha Excel rica em informações e visualizações, pronta para uso pela gestão escolar.

Este projeto demonstra a aplicação prática de Python para automação de tarefas administrativas e análise de dados no contexto educacional.

## 🚀 Recursos Principais

-   📄 **Geração Automatizada:** Cria planilhas de notas completas por disciplina e turma.
-   📊 **Dashboards Integrados:** Visualização de dados educacionais diretamente nas planilhas (desempenho, aprovação, evasão).
-   📈 **Análise de Desempenho:** Métricas por turma, disciplina e aluno individualmente.
-   🚦 **Controle de Situação Acadêmica:** Monitoramento de alunos (Ativos, Transferidos, Desistentes).
-   ⚙️ **Cálculos Automáticos:** Médias bimestrais/finais, taxas de aprovação/reprovação, e outros indicadores educacionais.
-   🎨 **Formatação Profissional:** Planilhas com layout claro, cores padronizadas, e logotipo institucional.
-   🔧 **Alta Configurabilidade:** Definição de disciplinas, estilos, fórmulas e estruturas via arquivos de configuração (`config.py` e JSON).

## 🛠️ Tecnologias Utilizadas

*   [![Python][Python-shield]][Python-url]
*   [![OpenPyXL][OpenPyXL-shield]][OpenPyXL-url]
*   Módulo `logging` (Python Standard Library)
*   Módulo `pathlib` (Python Standard Library)
*   Módulo `json` (Python Standard Library)

## 🖼️ Screenshots / Demonstração

<!-- IMPORTANTE: Adicione aqui screenshots das planilhas geradas! -->
<!-- Exemplo: -->
<!-- ![Dashboard Exemplo](link/para/sua/imagem_dashboard.png) -->
<!-- ![Planilha Disciplina](link/para/sua/imagem_planilha.png) -->
*Adicione aqui capturas de tela mostrando as diferentes abas da planilha, os dashboards e a formatação.*

## 📊 Estrutura da Planilha Gerada

O sistema gera um arquivo Excel (`.xlsx`) com uma estrutura organizada em múltiplas abas:

1.  **Abas por Disciplina:** (ex: Matemática, Português, etc.)
    *   Lista de alunos da turma.
    *   Colunas para notas bimestrais.
    *   Cálculo automático de médias.
    *   Coluna de Situação Final (Aprovado/Reprovado).
    *   *Dashboard* visual com gráficos de desempenho da turma na disciplina.
2.  **Aba SEC (Secretaria):**
    *   Coluna para Status do Aluno (Ativo, Transferido, Desistente).
    *   *Dashboards* com análise de evasão e taxas de aprovação gerais da turma.
3.  **Aba Boletim Consolidado:**
    *   Visão geral das médias e situação final de cada aluno em *todas* as disciplinas.
4.  **Abas Adicionais (Opcional/Configurável):**
    *   Relatórios individuais por aluno.
    *   Controle de Frequência.

## ⚙️ Configuração

A personalização do sistema é feita principalmente através de:

1.  **`config.py` (ou similar):**
    *   Definição da lista de disciplinas.
    *   Configuração de cores, fontes e estilos visuais.
    *   Ajuste fino das fórmulas de cálculo (se necessário).
    *   Definição da estrutura dos relatórios.
2.  **Arquivos JSON:**
    *   Armazenamento dos dados de entrada dos alunos (nomes, notas, status, etc.). É necessário preparar esses arquivos antes de executar o sistema.

## 📈 Indicadores Calculados

O sistema fornece automaticamente diversos indicadores chave:

*   Taxas de Aprovação e Reprovação (por turma e disciplina).
*   Médias Bimestrais e Finais.
*   Percentual de alunos com desempenho acima/abaixo da média.
*   Índices de Evasão (baseado no status Transferido/Desistente).
*   Situação Acadêmica final de cada aluno.

## 🚀 Como Executar

Siga os passos abaixo para configurar e executar o projeto:

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/lmbernardo7520112/teste-planilha-siage-interno.git
    cd teste-planilha-siage-interno
    ```
2.  **Crie um ambiente virtual (Recomendado):**
    ```bash
    python -m venv venv
    source venv/bin/activate  # Linux/macOS
    # venv\Scripts\activate  # Windows
    ```
3.  **Instale as dependências:**
    ```bash
    pip install openpyxl
    # Adicione outras dependências se houver um requirements.txt
    # pip install -r requirements.txt
    ```
4.  **Prepare os Dados:**
    *   Certifique-se de que os arquivos JSON com os dados dos alunos (notas, nomes, status) estão no local esperado pelo script e formatados corretamente.
    *   Revise e ajuste o arquivo `config.py` (ou similar) conforme necessário (disciplinas, nomes de turmas, etc.).
5.  **Execute o Script Principal:**
    ```bash
    python nome_do_script_principal.py
    ```
    *Substitua `nome_do_script_principal.py` pelo nome real do seu script principal.*

O script processará os dados e gerará o arquivo Excel na pasta de saída configurada.

## 🤝 Contribuição

Contribuições são bem-vindas! Se você tem sugestões para melhorar o sistema, sinta-se à vontade para:

1.  Fazer um Fork do projeto.
2.  Criar uma Branch para sua Feature (`git checkout -b feature/FuncionalidadeIncrivel`).
3.  Fazer Commit de suas alterações (`git commit -m 'Adiciona FuncionalidadeIncrivel'`).
4.  Fazer Push para a Branch (`git push origin feature/FuncionalidadeIncrivel`).
5.  Abrir um Pull Request.

Por favor, leia o `CONTRIBUTING.md` (se existir) para mais detalhes sobre o processo.

## 📜 Licença

Distribuído sob a licença MIT License. Veja `LICENSE` para mais informações.

<!-- CONTATOS -->
## 📧 Contato

 - [https://github.com/lmbernardo7520112](https://github.com/lmbernardo7520112) - lmbernardo752011@gmail.com

Link do Projeto: [https://github.com/lmbernardo7520112/teste-planilha-siage-interno](https://github.com/lmbernardo7520112/teste-planilha-siage-interno)

<!-- MARKDOWN LINKS & IMAGES -->
<!-- Corrija os links conforme necessário, especialmente para o arquivo LICENSE -->
[License-shield]: https://img.shields.io/github/license/lmbernardo7520112/teste-planilha-siage-interno?style=flat-square&color=informational
[License-url]: https://github.com/lmbernardo7520112/teste-planilha-siage-interno/blob/main/LICENSE
[Contributors-shield]: https://img.shields.io/github/contributors/lmbernardo7520112/teste-planilha-siage-interno?style=flat-square&color=informational
[Contributors-url]: https://github.com/lmbernardo7520112/teste-planilha-siage-interno/graphs/contributors
[Forks-shield]: https://img.shields.io/github/forks/lmbernardo7520112/teste-planilha-siage-interno?style=flat-square&color=informational
[Forks-url]: https://github.com/lmbernardo7520112/teste-planilha-siage-interno/network/members
[Stars-shield]: https://img.shields.io/github/stars/lmbernardo7520112/teste-planilha-siage-interno?style=flat-square&color=informational
[Stars-url]: https://github.com/lmbernardo7520112/teste-planilha-siage-interno/stargazers
[Issues-shield]: https://img.shields.io/github/issues/lmbernardo7520112/teste-planilha-siage-interno?style=flat-square&color=informational
[Issues-url]: https://github.com/lmbernardo7520112/teste-planilha-siage-interno/issues

[Python-shield]: https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white
[Python-url]: https://www.python.org/
[OpenPyXL-shield]: https://img.shields.io/badge/OpenPyXL-107C41?style=flat-square&logo=python&logoColor=white
[OpenPyXL-url]: https://openpyxl.readthedocs.io/en/stable/
