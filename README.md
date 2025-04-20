SIAGE INTERNO ECI LUIS RAMALHO - Sistema Integrado de Análise e Gestão Escolar


Sistema avançado para geração automatizada de planilhas de notas e relatórios escolares detalhados, projetado para otimizar processos e análises em instituições educacionais como a ECI Luis Ramalho.
✨ Sobre o Projeto
O SIAGE INTERNO foi desenvolvido para simplificar e automatizar a complexa tarefa de compilar notas, calcular médias, analisar o desempenho dos alunos e gerar relatórios consolidados. Utilizando Python e a biblioteca OpenPyXL, o sistema processa dados de entrada (no formato JSON) e produz uma planilha Excel rica em informações e visualizações, pronta para uso pela gestão escolar.
Este projeto demonstra a aplicação prática de Python para automação de tarefas administrativas e análise de dados no contexto educacional.
🚀 Recursos Principais

📄 Geração Automatizada: Cria planilhas de notas completas por disciplina e turma.
📊 Dashboards Integrados: Visualização de dados educacionais diretamente nas planilhas (desempenho, aprovação, evasão).
📈 Análise de Desempenho: Métricas por turma, disciplina e aluno individualmente.
🚦 Controle de Situação Acadêmica: Monitoramento de alunos (Ativos, Transferidos, Desistentes).
⚙️ Cálculos Automáticos: Médias bimestrais/finais, taxas de aprovação/reprovação, e outros indicadores educacionais.
🎨 Formatação Profissional: Planilhas com layout claro, cores padronizadas, e logotipo institucional.
🔧 Alta Configurabilidade: Definição de disciplinas, estilos, fórmulas e estruturas via arquivos de configuração (config.py e JSON).

🛠️ Tecnologias Utilizadas

 (versão 3.8 ou superior recomendada)

Módulo logging (Python Standard Library)
Módulo pathlib (Python Standard Library)
Módulo json (Python Standard Library)

🖼️ Screenshots / Demonstração




Em breve, serão adicionadas capturas de tela mostrando as diferentes abas da planilha, os dashboards e a formatação. (work in progress...)
📊 Estrutura da Planilha Gerada
O sistema gera um arquivo Excel (.xlsx) com uma estrutura organizada em múltiplas abas:

Abas por Disciplina: (ex: Matemática, Português, etc.)
Lista de alunos da turma.
Colunas para notas bimestrais (1º ao 4º bimestre).
Cálculo automático de médias (usando fórmulas Excel).
Coluna de Situação Final (Aprovado/Reprovado, baseado em média ≥ 7.0, por exemplo).
Dashboard visual com gráficos de desempenho da turma na disciplina.


Aba SEC (Secretaria):
Coluna para Status do Aluno (Ativo, Transferido, Desistente).
Dashboards com análise de evasão e taxas de aprovação gerais da turma.


Aba Boletim Consolidado:
Visão geral das médias e situação final de cada aluno em todas as disciplinas.


Abas Adicionais (Opcional/Configurável):
Relatórios individuais por aluno.
Controle de Frequência (se configurado no config.py).



⚙️ Configuração
A personalização do sistema é feita principalmente através de:

config.py (ou similar):

Lista de Disciplinas: Ex.: DISCIPLINES = ["Matemática", "Português", "Ciências"].
Estilos Visuais: Cores, fontes e bordas (usando OpenPyXL). Ex.: HEADER_COLOR = "FF0000" (vermelho para cabeçalhos).
Fórmulas de Cálculo: Critérios de aprovação. Ex.: APPROVAL_THRESHOLD = 7.0.
Estrutura das Abas: Quais abas incluir (ex.: incluir aba de frequência?).
Caminho dos Arquivos: Caminho do arquivo JSON de entrada (ex.: INPUT_PATH = "data/dados_turma.json").


Arquivos JSON:

Os dados dos alunos devem ser fornecidos em um arquivo JSON com a seguinte estrutura:{
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
Certifique-se de que as disciplinas no JSON correspondem às definidas no config.py.



📈 Indicadores Calculados
O sistema fornece automaticamente diversos indicadores chave:

Taxas de Aprovação e Reprovação (por turma e disciplina).
Médias Bimestrais e Finais (por aluno e disciplina).
Percentual de alunos com desempenho acima/abaixo da média da turma.
Índices de Evasão (baseado no status Transferido/Desistente).
Situação Acadêmica final de cada aluno (Aprovado/Reprovado com base na média configurada).

🚀 Como Executar
Siga os passos abaixo para configurar e executar o projeto:

Clone o repositório:git clone https://github.com/lmbernardo7520112/teste-planilha-siage-interno.git
cd teste-planilha-siage-interno


Crie um ambiente virtual (Recomendado):python -m venv venv
source venv/bin/activate  # Linux/macOS
venv\Scripts\activate     # Windows


Instale as dependências:pip install openpyxl

Nota: Caso exista um arquivo requirements.txt, use:pip install -r requirements.txt


Prepare os Dados:
Coloque o arquivo JSON de entrada (ex.: dados_turma.json) na pasta data/.
Edite o arquivo config.py para ajustar:
Lista de disciplinas.
Caminho do arquivo de entrada (se diferente de data/dados_turma.json).
Critérios de aprovação e estilos visuais.




Execute o Script Principal:O script principal é gerar_planilha.py. Execute:python gerar_planilha.py

O script gerará o arquivo Excel na pasta output/ (ex.: output/relatorio_turma_9A.xlsx).

Possíveis Erros e Soluções

Erro: "FileNotFoundError: data/dados_turma.json"Verifique se o arquivo JSON está na pasta data/ e se o nome está correto no config.py.
Erro: "KeyError: 'Matemática'"Certifique-se de que todas as disciplinas no JSON correspondem às definidas em DISCIPLINES no config.py.
Erro: "ModuleNotFoundError: No module named 'openpyxl'"Instale a biblioteca OpenPyXL com pip install openpyxl.

🤝 Contribuição
Contribuições são bem-vindas! Se você tem sugestões para melhorar o sistema, sinta-se à vontade para:

Fazer um Fork do projeto.
Criar uma Branch para sua Feature (git checkout -b feature/FuncionalidadeIncrivel).
Fazer Commit de suas alterações (git commit -m 'Adiciona FuncionalidadeIncrivel').
Fazer Push para a Branch (git push origin feature/FuncionalidadeIncrivel).
Abrir um Pull Request.

Por favor, leia o CONTRIBUTING.md (se existir) para mais detalhes sobre o processo.
📜 Licença
Distribuído sob a licença MIT License. Veja LICENSE para mais informações.
📧 Contato

https://github.com/lmbernardo7520112 - lmbernardo752011@gmail.com

Link do Projeto: https://github.com/lmbernardo7520112/teste-planilha-siage-interno

