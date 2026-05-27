# Configurar recepção de pedidos contraentrega (Google Sheets)

Cada pedido enviado pelo formulário vai virar uma linha na sua planilha do Google Sheets. Setup leva ~5 minutos, é grátis e não depende de servidor.

## 1) Criar a planilha

1. Acesse <https://sheets.google.com> → **Em branco**
2. Renomeie a planilha (ex.: `Pedidos Elyza PT`)
3. Anote o nome da aba de baixo (por padrão `Hoja 1` ou `Sheet1`)

## 2) Colar o script

1. Na planilha: **Extensiones** → **Apps Script**
2. Apague o conteúdo padrão e cole o código abaixo
3. Se sua aba não se chama `Pedidos`, troque a string na linha do `SHEET_NAME`
4. Salve (💾 ou `Cmd+S`) e dê um nome ao projeto (ex.: `elyza-cod`)

```javascript
const SHEET_NAME = 'Pedidos'; // nome da aba — ajuste se necessário

const HEADERS = [
  'received_at','order_id','created_at',
  'name','phone','email',
  'address','city','zip','province','notes',
  'product','sku','quantity','unit_price','total','currency',
  'payment_method','country','url'
];

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);

    if (sh.getLastRow() === 0) {
      sh.appendRow(HEADERS);
      sh.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold').setBackground('#0d0d0d').setFontColor('#fff');
      sh.setFrozenRows(1);
    }

    const row = HEADERS.map(h => h === 'received_at' ? new Date() : (data[h] || ''));
    sh.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: String(err)}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput('Elyza COD endpoint OK');
}
```

## 3) Publicar como Web App

1. No editor do Apps Script: **Implementar** (botão azul, canto sup. direito) → **Nueva implementación**
2. Em **Tipo**, clique no ⚙️ engrenagem → **Aplicación web**
3. Configurar:
   - **Descripción**: `Elyza COD`
   - **Ejecutar como**: **Yo** (a tua conta)
   - **Quién tiene acceso**: **Cualquier persona** ← importante!
4. **Implementar** → o Google pedirá autorização (clique em **Autorizar acceso**, escolha sua conta, em "Google no ha verificado…" clique em **Configuración avanzada** → **Ir a [seu projeto] (no seguro)** → **Permitir**)
5. Copie a **URL de la aplicación web** (algo como `https://script.google.com/macros/s/AKfycb.../exec`)

## 4) Colar a URL na página

Abra `index.html` e procure por este bloco no topo (logo após o Meta Pixel):

```html
<script>
  window.ELYZA_COD_ENDPOINT = ""; // ← pega aquí tu URL de Apps Script
</script>
```

Cole sua URL entre as aspas:

```html
<script>
  window.ELYZA_COD_ENDPOINT = "https://script.google.com/macros/s/AKfycb.../exec";
</script>
```

Faça commit e push para o GitHub. **Pronto.**

## 5) Testar

1. Abra a sua página (no GitHub Pages ou localmente)
2. Clique em "Pedir contra reembolso", preencha o formulário, envie
3. Abra sua planilha → o pedido deve aparecer em segundos com timestamp

Se não aparecer:
- Verifique se em **Implementar > Administrar implementaciones** a versão está "Activa" e o acesso "Cualquier persona"
- No DevTools (F12) → aba **Console**, digite `elyzaOrders.list()` — vê todos os pedidos do navegador atual (backup local)

## 6) Cada nova alteração no script

Se editar o script, precisa criar uma **Nueva implementación** (ou usar **Probar implementaciones**) para que mude no ar. Não basta salvar.

## 7) Atualizar a URL no futuro

Se trocar a implantação, copie a URL nova e atualize o `window.ELYZA_COD_ENDPOINT` no `index.html`.

---

## Backup local (sempre ativo)

Independentemente do endpoint, **todos os pedidos também ficam salvos no navegador** do cliente (no `localStorage`). Isso protege contra erros de rede e quedas do Google. Para acessar os backups do seu próprio navegador (útil para testes), abra o DevTools (F12) na sua página e use no Console:

```js
elyzaOrders.list()      // mostra todos os pedidos salvos
elyzaOrders.count()     // quantos pedidos
elyzaOrders.csv()       // exporta como CSV (Excel/Sheets)
elyzaOrders.download()  // baixa um .csv
elyzaOrders.clear()     // apaga (cuidado)
```

⚠️ Esse backup só serve para os pedidos feitos **a partir do seu próprio navegador** — cada cliente que pede tem o backup só na máquina dele. Por isso o Google Sheets é a fonte de verdade.

---

## Receber notificação por email a cada pedido (opcional)

Na própria planilha: **Herramientas** → **Reglas de notificación** → **Cualquier cambio** → **Enviar email inmediatamente**.

Ou no Apps Script, adicione antes do `return` do `doPost`:

```javascript
MailApp.sendEmail({
  to: 'seu-email@gmail.com',
  subject: `Nuevo pedido COD — ${data.name} (€${data.total})`,
  body: `Cliente: ${data.name}\nTel: ${data.phone}\nDirección: ${data.address}, ${data.city} ${data.zip}, ${data.province}\nTotal: €${data.total}\nObservaciones: ${data.notes || '-'}`
});
```

Republique a implantação após editar.
