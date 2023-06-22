# Sistema de automatização de emissão de notas fiscais de serviço da Prefeitura de BH

# O Problema

>>Antes, todo o processo era realizado de forma manual, no qual incorria em vários erros, dado a rotina enfadonha de copia e cola, o que ocasionava em preenchimentos incorretos e, consequentemente, cancelamentos de documentos fiscais.
Ademais, a pessoa que realizava a operação era uma idosa que não possuía uma base consistente no que tange à tecnologia. 
Cada NF gerada podia levar 10 minutos, sem contar, evidentemente, o retrabalho quando de emissão incorreta.
Com a automatização, o tempo de emissão de cada NF caiu para em média 30 segundos!, além de proporcionar segurança dos dados por meio de critérios bem delineados.

## Desenvolvido em Python e com o auxílio das bibliotecas: Selenium e PyAutoGui

## Base de dados: Google Sheets

### Resumo da aplicação:

>O sistema coleta os dados numa planilha do Google Sheets, por meio da api disponibilizada pelo Google, armazena-os num dicionário Python e preenche os campos obrigatórios, com base nos dados do dicionário, da página de emissão de NF de serviço da prefeitura de BH.

### Google Sheets:

Será necessário ler a [documentação](https://developers.google.com/sheets/api/guides/concepts?hl=pt-br) e seguir os passos para acessar a api do Google Sheets e Google Drive.

Supondo que tenha conta no gmail e criado a planilha, o próximo passo é nomear as abas e popular a planilha com os dados exigidos para emissão de NF, isto é:
- Data Emissão
- Nome do Cliente
- CPF
- Endereço
- Número
- Bairro
- Município
- Valor

No meu caso existem vários outros campos, uma vez que são usados para outras finalidades que não emissão de NF. Todavia, o script só irá considerar os dados citados acima.

Para indicar para o script quais dados deverão ser coletados, optei em colocar a palavra "Liberado" na linha correspondente aos dados que pretendo que sejam coletados e que serão utilizados para a emissao da NF.

![Captura de tela de 2023-06-15 21-44-08](https://github.com/ch-soares/emissor-nf-prefeitura-bh/assets/65301099/8e028a10-7df8-4dac-8675-39a536026319)

>>Os dados são devidamente tratados por meio critérios, utilizando-se de fórmulas do google sheets, de modo a não permitir que informações inconsistentes sejam coletadas.

O projeto possui a seguinte estrutura:

```
.
├── chromedriver
│   └── chromedriver.exe
├── credenciais.py
├── .env
├── .gitignore
├── google_sheets
│   ├── __init__.py
│   ├── service_account.json
│   └── sheets.py
├── nf
│   ├── emissor_nf.py
│   └── nfs_emitidas.txt
├── README.md
└── requirements.txt

```

![Captura de tela de 2023-06-16 10-16-40.png](..%2F..%2FImagens%2FCaptura%20de%20tela%20de%202023-06-16%2010-16-40.png)

- O módulo credenciais.py é onde se dá a parametrização, por meio da biblioteca python-decouple, que de modo resumido gerencia variáveis de ambiente de forma a não permitir que dados sensíveis, como neste caso: login; XPATH dos elementos da página da prefeitura de BH; ID da planilha etc fiquem expostos.
Tais dados devem ser expressados num arquivo .env, que neste repositório está nomeado como .env-example, visto que o arquivo .env original não pode em nenhuma hipótese ser público, pelas razões já informadas.

- O arquivo .env, como já mencionado, é onde estão relacionadas todas as variáveis de ambiente, fundamentais para o funcionamento do sistema. Neste repositório, no entanto, este arquivo está nomeado como env-example, no qual estão descritas apenas os nomes das variáveis, sem os valores, a título de exemplo. 

- No pacote google_sheets, o módulo sheets.py é onde se faz a comunicação com o planilha. Esta comunicação é realizada por meio da biblioteca [Gspread](https://docs.gspread.org/en/v5.7.2/). As funções deste módulo são responsáveis por:

  - Coleta de dados na planilha;
  - Armazenagem dos dados das NFs emitidas;
  - E atualização da planilha com as Nfs emitidas

- O múdulo emissor_nf.py, dentro do pacote nf, é onde está toda a lógica por trás da emissão das NFs. Ele acessa os dados gerados do módulo sheets.py e por meio do Selenium e PyAutoGui os dados são preenchidos na plataforma da prefeitura de BH.
O script é capaz de gerar vários documentos com apenas um comando, por meio de um loop.

- O arquivo requirements.tx é onde estão relacionadas todas as dependências do projeto.
  
- Já o arquivo .gitignore é onde estão descritos todos os diretórios e arquivos que não precisam e/ou não podem ir para o repositório remoto, como: .env, .venv, service_account.json etc.

Nota-se que neste repositório não se encontram o diretório chromedriver nem o arquivo executável de mesmo nome, uma vez que tal arquivo muda de acordo com a versão do browser do Google.
Devendo, portanto, ser baixado de acordo com cada contexto.
 
O arquivo service.account.json também não se encontra neste repositório em razão deste arquivo conter dados sensíveis da planilha, como token e ID, por exemplo; portanto não podem ser públicos.
Este arquivo deverá ser devidamente baixado após ler a [documentação](https://developers.google.com/sheets/api/guides/concepts?hl=pt-br) para acessar a api do Google Sheets e executar os comandos impostos pela plataforma.

O arquivo nfs_emitidas.txt não está neste repositório em função de conter os dados do cliente cuja nf foi emitida. De igual sorte, portanto, não pode ser público. Optou-se por armazenar os dados de emissão e atualização da planilha num arquivo txt de modo a diminuir os acessos à api e evitar erros eventuais na navegação web, e que poderiam interromper a execução do script.
