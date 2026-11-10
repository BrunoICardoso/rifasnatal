# Rifa de Natal – como colocar no ar

## 1. Site (GitHub Pages)
Os arquivos já estão neste repositório. Para publicar: *Settings → Pages → Deploy from a branch → main / (root) → Save*.
O endereço fica https://brunoicardoso.github.io/rifasnatal/

## 2. Controle automático (planilha + Apps Script), uns 5 minutos
A planilha já existe no Drive, na pasta da rifa: **Rifa de Natal 2026 – vendas**
(https://docs.google.com/spreadsheets/d/1wv20MKrJ1ehDy_LNn3aIgwL4iGOB76jlVSNgOqXp4vs/edit).

1. Abre a planilha e vai em *Extensões → Apps Script*. Apaga o que tiver e cola o conteúdo de `apps-script/Code.gs`. Salva.
2. No topo, escolhe a função **configurar** e clica em **Executar**. Autoriza com a tua conta
   (vai aparecer "app não verificado": *Avançado → Acessar*). Isso aplica a formatação e os menus da coluna Situação.
3. **Implantar → Nova implantação →** tipo **App da Web**
   - Executar como: **Eu**
   - Quem pode acessar: **Qualquer pessoa**

   Copia a URL que termina em `/exec`.
4. No `index.html`, cola essa URL em `"apiUrl": ""` e faz o commit.

Se mudar o `Code.gs` depois: *Implantar → Gerenciar implantações → editar (lápis) → Nova versão*,
para manter a mesma URL.

## 3. Como a organização controla
Cada compra vira uma linha na aba **Vendas** sozinha, e a pessoa manda o comprovante no WhatsApp.
A organização só precisa olhar a coluna **Situação**:

| Situação | O que significa | O que fazer |
|---|---|---|
| Aguardando pagamento | A pessoa está pagando. Se em 30 min não disser que pagou, os números voltam a ficar livres sozinhos. | Nada |
| Pago - conferir | A pessoa disse que pagou e mandou o comprovante no WhatsApp. | Conferir o PIX no extrato e trocar para **Confirmado** |
| Confirmado | PIX conferido. | Nada |
| Cancelado | Libera os números. | Usar quando o PIX não veio |
| Pago - CONFLITO | Pagou depois do prazo e alguém já tinha pego algum número. | Combinar com a pessoa: devolver ou trocar o número |

A aba **Resumo** mostra o total vendido, o valor confirmado e o que falta conferir.
O site mostra só o primeiro nome. Telefones ficam só na planilha, que não deve ser compartilhada publicamente.

## 4. Depois do sorteio
Preenche `"resultado"` no `index.html`, por exemplo:
`[{"premio":"1º prêmio","numero":427},{"premio":"2º prêmio","numero":81},{"premio":"3º prêmio","numero":930}]`.
O site mostra os ganhadores e para de vender.
