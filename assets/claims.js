/* MAGNIT RC — Grand-Trade claims text constructor. No server, no persistent storage. */
(() => {
  'use strict';
  const $ = id => document.getElementById(id);
  const form = $('claimForm');
  if (!form) return;
  const itemsHost = $('claimItems');
  const output = $('claimOutput');
  const status = $('claimFormStatus');
  const copyButton = $('claimCopy');
  const saveButton = $('claimSaveTxt');
  const countNode = $('claimFieldCount');
  let itemSequence = 0;
  let finalText = '';

  function itemTemplate(id) {
    const node = document.createElement('div');
    node.className = 'claim-item';
    node.dataset.itemId = String(id);
    node.innerHTML = `
      <div class="claim-item-head"><h3 class="claim-item-title">Товарная позиция</h3><button type="button" class="claim-remove" aria-label="Удалить эту товарную позицию">Удалить</button></div>
      <div class="claims-fields">
        <label class="field claim-wide">Наименование товарной позиции <span class="req">*</span><input data-item-field="name" type="text" required autocomplete="off" maxlength="400" placeholder="Полное наименование из документов"></label>
        <label class="field">Фактическое количество к возврату <span class="req">*</span><input data-item-field="qty" type="text" inputmode="decimal" required autocomplete="off" maxlength="80" placeholder="Количество и единица измерения"></label>
        <label class="field">Сколько палет <span class="req">*</span><input data-item-field="pallets" type="text" inputmode="decimal" required autocomplete="off" maxlength="80" placeholder="Количество палет"></label>
        <label class="field claim-wide">Причина <span class="req">*</span><textarea data-item-field="reason" rows="3" required maxlength="1000" placeholder="Укажите фактическую причину возврата"></textarea></label>
      </div>`;
    node.querySelector('.claim-remove').addEventListener('click', () => {
      if (itemsHost.children.length < 2) return;
      node.remove();
      refreshItemNumbers();
      refreshPreview();
    });
    node.addEventListener('input', refreshPreview);
    return node;
  }
  function refreshItemNumbers() {
    [...itemsHost.children].forEach((node, index) => {
      node.querySelector('.claim-item-title').textContent = `Товарная позиция ${index + 1}`;
      node.querySelector('.claim-remove').hidden = itemsHost.children.length === 1;
    });
  }
  function addItem() {
    const node = itemTemplate(++itemSequence);
    itemsHost.append(node);
    refreshItemNumbers();
    refreshPreview();
    return node;
  }
  function getItems() {
    return [...itemsHost.children].map(node => {
      const get = key => node.querySelector(`[data-item-field="${key}"]`).value.trim();
      return { name:get('name'), qty:get('qty'), pallets:get('pallets'), reason:get('reason') };
    });
  }
  const get = id => $(id).value.trim();
  function orderNumber(raw) {
    if (!raw) return '';
    return /^ГТАдг/i.test(raw) ? raw : `ГТАдг${raw}`;
  }
  function claimData() {
    return {
      vehicle:get('claimVehicle'),
      order:orderNumber(get('claimOrder')),
      request:get('claimRequest'),
      gu:get('claimGU'),
      gt:get('claimGT'),
      items:getItems()
    };
  }
  function buildText(data, placeholder = false) {
    const val = v => v || (placeholder ? '—' : '');
    const lines = [
      'ПРЕТЕНЗИЯ — ГРАНД-ТРЕЙД',
      '',
      'Контрагент: Гранд-Трейд',
      `Номер автомобиля: ${val(data.vehicle)}`,
      `Номер заказа: ${val(data.order)}`,
      `Номер заявки: ${val(data.request)}`,
      `GU: ${val(data.gu)}`,
      `GT: ${val(data.gt)}`,
      '',
      'ТОВАРНЫЕ ПОЗИЦИИ К ВОЗВРАТУ'
    ];
    data.items.forEach((item, index) => {
      lines.push(
        '',
        `${index+1}. Наименование товарной позиции: ${val(item.name)}`,
        `   Фактическое количество к возврату: ${val(item.qty)}`,
        `   Количество палет: ${val(item.pallets)}`,
        `   Причина: ${val(item.reason)}`
      );
    });
    return lines.join('\n');
  }
  function requiredEntries() {
    return [
      ...['claimVehicle','claimOrder','claimRequest','claimGU','claimGT'].map(id => $(id)),
      ...itemsHost.querySelectorAll('input[required], textarea[required]')
    ];
  }
  function refreshPreview() {
    const entries=requiredEntries();
    const filled=entries.filter(node => node.value.trim()).length;
    countNode.textContent=`${filled} / ${entries.length} полей`;
    output.value=buildText(claimData(),true);
    finalText='';
    copyButton.disabled=true;
    saveButton.disabled=true;
    status.textContent=filled===entries.length?'Все поля заполнены. Нажмите «Сформировать текст».':`Заполнено ${filled} из ${entries.length} обязательных полей.`;
    status.classList.remove('claim-success');
  }
  $('claimAddItem').addEventListener('click', () => {const item = addItem(); item.querySelector('[data-item-field="name"]').focus();});
  form.addEventListener('input', refreshPreview);
  form.addEventListener('submit', e => {
    e.preventDefault();
    const missing=requiredEntries().find(node => !node.value.trim());
    if (missing) {
      status.textContent='Заполните обязательные поля, отмеченные звёздочкой.';
      status.classList.remove('claim-success');
      missing.focus();
      return;
    }
    finalText=buildText(claimData());
    output.value=finalText;
    copyButton.disabled=false;
    saveButton.disabled=false;
    status.textContent='Текст сформирован. Можно скопировать или сохранить .txt.';
    status.classList.add('claim-success');
    if(window.matchMedia('(max-width: 920px)').matches) output.scrollIntoView({behavior:'smooth',block:'center'});
  });
  copyButton.addEventListener('click', async () => {
    if (!finalText) return;
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(finalText);
      } else {
        output.focus();output.select();
        if (!document.execCommand('copy')) throw new Error('copy failed');
      }
      status.textContent='Текст скопирован в буфер обмена.';
      status.classList.add('claim-success');
    } catch (_) {
      output.focus();output.select();
      status.textContent='Не удалось скопировать автоматически. Текст выделен — скопируйте его вручную.';
    }
  });
  saveButton.addEventListener('click', () => {
    if (!finalText) return;
    const blob=new Blob(['\ufeff',finalText],{type:'text/plain;charset=utf-8'});
    const url=URL.createObjectURL(blob);
    const anchor=document.createElement('a');
    anchor.href=url;
    anchor.download='Pretenziya_Grand_Trade.txt';
    document.body.append(anchor);
    anchor.click();anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 120000);
    status.textContent='Текстовый файл подготовлен к скачиванию.';
  });
  $('claimClear').addEventListener('click', () => {
    form.reset();
    itemsHost.replaceChildren();
    addItem();
    refreshPreview();
    $('claimVehicle').focus();
  });
  addItem();
})();
