SIAGE INTERNO ECI LUIS RAMALHO - Sistema Integrado de Análise e Gestão Escolar


Sistema avançado para geração automatizada de planilhas de notas e relatórios escolares detalhados, projetado para otimizar processos e análises em instituições educacionais como a ECI Luis Ramalho.
✨ Sobre o Projeto
O SIAGE INTERNO foi desenvolvido para simplificar e automatizar a tarefa de compilar notas, calcular médias, analisar o desempenho dos alunos e gerar relatórios consolidados para a gestão escolar. Utilizando Python e a biblioteca OpenPyXL, o sistema processa dados de entrada no formato JSON e produz uma planilha Excel (.xlsx) com informações detalhadas, incluindo dashboards visuais e cálculos automáticos, pronta para uso pela equipe administrativa.
Este projeto é uma aplicação prática de automação e análise de dados educacionais, voltada para facilitar o dia a dia de escolas e secretarias educacionais.
🚀 Recursos Principais

📄 Geração Automatizada: Cria planilhas de notas completas por disciplina e turma.
📊 Dashboards Integrados: Visualização de dados educacionais diretamente nas planilhas (desempenho, aprovação, evasão).
📈 Análise de Desempenho: Métricas por turma, disciplina e aluno individualmente.
🚦 Controle de Situação Acadêmica: Monitoramento de alunos (Ativos, Transferidos, Desistentes).
⚙️ Cálculos Automáticos: Médias bimestrais/finais, taxas de aprovação/reprovação, e outros indicadores educacionais.
🎨 Formatação Profissional: Planilhas com layout claro, cores padronizadas, e logotipo institucional.
🔧 Alta Configurabilidade: Definição de disciplinas, estilos, fórmulas e estruturas via config.py e arquivos JSON.

🛠️ Tecnologias Utilizadas

 (versão 3.8 ou superior recomendada)
 (para manipulação de planilhas Excel)
Módulo logging (Python Standard Library) - Para rastreamento de erros e logs
Módulo pathlib (Python Standard Library) - Para manipulação de caminhos de arquivos
Módulo json (Python Standard Library) - Para leitura dos dados de entrada

🖼️ Screenshots / Demonstração
Em breve, serão adicionadas capturas de tela mostrando:

A aba de uma disciplina com notas e médias.
O dashboard de desempenho da turma.
A aba SEC com análise de evasão e taxas de aprovação.

📊 Estrutura da Planilha Gerada
O sistema gera um arquivo Excel (.xlsx) com a seguinte estrutura:

Abas por Disciplina (ex: Matemática, Português):

Lista de alunos da turma.
Colunas para notas bimestrais (1º ao 4º bimestre).
Cálculo automático de médias (usando fórmulas Excel).
Coluna de Situação Final (Aprovado/Reprovado, baseado em média ≥ 7.0, por exemplo).
Gráficos de desempenho da turma na disciplina.


Aba SEC (Secretaria):

Coluna para Status do Aluno (Ativo, Transferido, Desistente).
Dashboards com análise de evasão e taxas de aprovação gerais da turma.


Aba Boletim Consolidado:

Visão geral das médias e situação final de cada aluno em todas as disciplinas.


Abas Adicionais (Configuráveis):

Relatórios individuais por aluno.
Controle de frequência (se configurado no config.py).



⚙️ Configuração
O sistema é altamente configurável por meio de dois componentes principais:
1. config.py
Este arquivo contém as configurações principais do sistema. Exemplos de configurações:

Lista de Disciplinas: Ex.: DISCIPLINES = ["Matemática", "Português", "Ciências"]
Estilos Visuais: Cores das células, fontes e bordas (usando OpenPyXL).
Ex.: HEADER_COLOR = "FF0000" (vermelho para cabeçalhos).


Fórmulas de Cálculo: Critérios de aprovação (ex.: média mínima).
Ex.: APPROVAL_THRESHOLD = 7.0


Estrutura das Abas: Quais abas incluir (ex.: incluir aba de frequência?).

2. Arquivos JSON de Entrada
Os dados dos alunos devem ser fornecidos em um arquivo JSON com a seguinte estrutura:
{
  "turma": "9A",
  "alunos": [
    {
      "nome": "João Silva",
      "status": "Ativo",
      "notas": {
        "Matemática": [8.5, 7.0, 9.0, 6.5],
        "Português": [6.0, 5.5, 7.0, 8.0]
      }
    },
    {
      "nome": "Maria Oliveira",
      "status": "Transferido",
      "notas": {
        "Matemática": [5.0, 4.5, 6.0, 5.5],
        "Português": [7.0, 6.5, 8.0, 7.5]
      }
    }
  ]
}


O arquivo deve estar na pasta data/ (ou conforme configurado no config.py).
Certifique-se de que todas as disciplinas listadas no JSON correspondem às definidas no config.py.

📈 Indicadores Calculados
O sistema calcula automaticamente:

Taxas de Aprovação e Reprovação (por turma e disciplina).
Médias Bimestrais e Finais (por aluno e disciplina).
Percentual de alunos com desempenho acima/abaixo da média da turma.
Índices de Evasão (calculado com base nos status "Transferido" e "Desistente").
Situação Acadêmica Final (Aprovado/Reprovado com base na média configurada).

🚀 Como Executar
Siga os passos abaixo para configurar e executar o projeto:
1. Clone o Repositório
git clone https://github.com/lmbernardo7520112/teste-planilha-siage-interno.git
cd teste-planilha-siage-interno

2. Crie um Ambiente Virtual (Recomendado)
python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows

3. Instale as Dependências
pip install openpyxl

Nota: Caso exista um arquivo requirements.txt, use:
pip install -r requirements.txt

4. Prepare os Dados

Coloque o arquivo JSON de entrada (ex.: dados_turma.json) na pasta data/.
Edite o arquivo config.py para ajustar:
Lista de disciplinas.
Caminho do arquivo de entrada (se diferente de data/dados_turma.json).
Critérios de aprovação e estilos visuais.



5. Execute o Script Principal
O script principal é gerar_planilha.py. Execute:
python gerar_planilha.py


O script lerá os dados do JSON, processará as notas e gerará o arquivo Excel na pasta output/ (ex.: output/relatorio_turma_9A.xlsx).

Possíveis Erros e Soluções

Erro: "FileNotFoundError: data/dados_turma.json"
Verifique se o arquivo JSON está na pasta correta (data/) e se o nome está correto no config.py.


Erro: "KeyError: 'Matemática'"
Certifique-se de que todas as disciplinas no JSON correspondem às definidas em DISCIPLINES no config.py.


Erro: "ModuleNotFoundError: No module named 'openpyxl'"
Instale a biblioteca OpenPyXL com pip install openpyxl.



🤝 Contribuição
Contribuições são bem-vindas! Para contribuir:

Faça um Fork do projeto.
Crie uma Branch para sua feature (git checkout -b feature/NovaFuncionalidade).
Faça Commit das alterações (git commit -m 'Adiciona NovaFuncionalidade').
Faça Push para a Branch (git push origin feature/NovaFuncionalidade).
Abra um Pull Request.

📜 Licença
Distribuído sob a licença MIT. Veja o arquivo LICENSE para mais informações.
📧 Contato

https://github.com/lmbernardo7520112 - lmbernardo752011@gmail.com

Link do Projeto: https://github.com/lmbernardo7520112/teste-planilha-siage-interno

