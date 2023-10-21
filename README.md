## Processamento de vendas
### CASO REAL

No momento em que escrevo esse código, presto serviços pra uma determinada empresa que precisa de dados de vendas feitas no
ecommerce e processadas por um determinado gateway de pagamentos.

Acontece que no ecommerce não tem todos os dados necessários, como o código de autorização da venda e a data de previsão
para receber o primeiro valor.

Nesse caso, eu:
- Puxo os dados da venda do nosso site via API do WooCommerce;
- Extraio e transformo os dados de um relatório que exporto em formato '.xls' do gateway de pagamentos porque ele não
possui API disponível;
- Unifico os dados relevantes baseado num identificador que existe tanto no ecommerce quanto no gateway (nsu ou _fixpay_tid);
- Exporto os dados para excel.

OBS: Esse código processa apenas vendas feitas no cartão e recebidas no gateway.
