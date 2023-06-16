# Sistema de automatização de emissão de notas fiscais de serviço da Prefeitura de BH

## Desenvolvido em Python e com o auxílio das bibliotecas: Selenium e PyAutoGui

## Base de dados: Google Sheets

### Resumo da aplicação:

>O sistema coleta os dados numa planilha do Google Sheets, armazena-os num dicionário Python e preenche os campos obrigatórios, com base nos dados do dicionário, da página de emissão de NF de serviço da prefeitura de BH.

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

No meu caso existem vários outros campos, uma vez que são usados para outras finalidades que não emissão de NF. Desta forma, se desejar, você poderá fazê-lo também, pois o script só irá considerar os dados citados acima.

Para indicar para o script quais dados deverão ser coletados, optei em colocar a palavra "Liberado" na linha correspondente aos dados que pretendo que sejam coletados e que serão utilizados para a emissao da NF.

>>Estes dados são devidamente tratados por meio critérios, utilizando-se de fórmulas do google sheets, de modo a não permitir que dados inconsistentes sejam coletados.

O projeto possui a seguinte estrutura:

```
.
├── chromedriver
│   └── chromedriver.exe
├── credenciais.py
├── google_sheets
│   ├── __init__.py
│   ├── service_account.json
│   └── sheets.py
├── nf
│   └── emissor_nf.py
├── README.md
└── requirements.txt

```

- O módulo credenciais.py é onde se dá a parametrização, por meio da biblioteca python-decouple, que de modo resumido gerencia variáveis de ambiente de forma a não permitir que dados sensíveis, como neste caso: login; XPATH dos elementos da página da prefeitura de BH; ID da planilha etc fiquem expostos.
Tais dados devem ser expressados num arquivo .env, que neste repositório está nomeado como .env-example.

- No pacote google_sheets, o módulo sheets.py possui funções que são responsáveis por:

  - Coleta de dados na planilha;
  - Armazenagem dos dados das NFs emitidas;
  - E atualização da planilha com as Nfs emitidas

- O múdulo emissor_nf.py, dentro do pacote nf, é onde está toda a lógica por trás da emissão das NFs. Ele acessa os dados gerados do módulo sheets.py e por meio do Selenium e PyAutoGui os dados são preenchidos na plataforma da prefeitura de BH.
O script é capaz de gerar vários documentos com apenas um comando, por meio de um loop.

- O arquivo requirements.tx é onde estão relacionadas todas as dependências do projeto.

Nota-se que neste repositório não se encontram o diretório chromedriver nem o arquivo executável de mesmo nome, uma vez que tal arquivo muda de acordo com a versão do browser do Google.
Devendo, portanto, ser baixado de acordo com cada contexto.
 
O arquivo service.account.json também não se encontra neste repositório em razão deste arquivo conter dados sensíveis da planilha, como token e ID, por exemplo; portanto não podem ser públicos.
Este arquivo deverá ser devidamente baixado após ler a [documentação](https://developers.google.com/sheets/api/guides/concepts?hl=pt-br) para acessar a api do Google Sheets e executar os comandos impostos pela plataforma.

Optou-se por armazenar os dados de emissão e atualização da planilha num arquivo txt de modo a diminuir os acessos à api e evitar erros eventuais na navegação web, que poderiam interromper a execução do script.

### O Problema

>>Antes, todo o processo era realizado de forma manual, no qual incorria em vários erros, dado a rotina enfadonha de copia e cola, o que ocasionava em preenchimentos incorretos e, consequentemente, cancelamentos de documentos fiscais.
Ademais, a pessoa que realizava a operação era uma idosa que não possuía uma base consistente de questões relacionadas à tecnologia. 
Cada NF gerada podia levar 10 minutos, sem contar, evidentemente, o retrabalho quando de emissão incorreta.
Com a automatização, o tempo de emissão de cada NF é em média 30 segundos!

