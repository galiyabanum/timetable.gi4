(function(root,factory){
  const api=factory(root.XLSX||(typeof require==='function'?require('xlsx'):null));
  if(typeof module==='object'&&module.exports)module.exports=api;
  root.TimetableImporter=api;
})(typeof globalThis!=='undefined'?globalThis:this,function(XLSX){
  'use strict';
  const DAYS=['Пн','Вт','Ср','Чт','Пт','Сб'];
  const text=value=>value==null?'':String(value).replace(/\u00a0/g,' ').replace(/[ \t]+/g,' ').trim();
  const tokens=value=>text(value).split(/\r?\n/).map(text).filter(Boolean);
  const isDash=value=>/^[-–—]+$/.test(text(value));
  const className=value=>{
    const match=text(value).match(/^Класс\s*[-–—:]\s*(.+)$/i);
    return match?text(match[1]).replace(/(\d)\s*([А-ЯA-Z])$/u,'$1 $2'):'';
  };
  const roomName=value=>{
    const clean=text(value);
    if(!clean||isDash(clean))return '';
    const number=clean.match(/^([0-9]+(?:[А-ЯA-ZА-Яа-я-]*)?)/u);
    return number?number[1]:clean;
  };
  function rowsForSheet(workbook,name){
    return XLSX.utils.sheet_to_json(workbook.Sheets[name],{header:1,defval:null,raw:false,blankrows:true});
  }
  function findSheet(workbook,part){
    return workbook.SheetNames.find(name=>text(name).toLocaleLowerCase('ru').includes(part));
  }
  function teacherNames(workbook){
    const name=findSheet(workbook,'учител');
    if(!name)return new Set();
    const rows=rowsForSheet(workbook,name),result=new Set();
    for(let row=0;row<rows.length;row++){
      const candidate=text(rows[row]?.[0]);
      const header=text(rows[row+2]?.[0]);
      if(candidate&&header==='#')result.add(candidate);
    }
    return result;
  }
  function looksLikeTeacher(value,known){
    const clean=text(value);
    return known.has(clean)||/^[А-ЯЁ][а-яё-]+(?:\s+[А-ЯЁ]\.){1,2}$/u.test(clean);
  }
  function parse(buffer){
    if(!XLSX)throw new Error('Модуль чтения Excel не загрузился. Обновите страницу и попробуйте снова.');
    let workbook;
    try{workbook=XLSX.read(buffer,{type:'array',cellDates:false})}catch(error){throw new Error('Не удалось открыть файл Excel. Возможно, файл повреждён или защищён паролем.')}
    const classSheet=findSheet(workbook,'класс');
    if(!classSheet)throw new Error('В файле не найден лист с классами. Название листа должно содержать слово «классы».');
    const rows=rowsForSheet(workbook,classSheet),knownTeachers=teacherNames(workbook);
    const blocks=[];
    for(let row=0;row<rows.length;row++){
      const name=className(rows[row]?.[0]);
      if(name)blocks.push({row,name});
    }
    if(!blocks.length)throw new Error('Не найдены разделы вида «Класс - 5 А». Проверьте структуру файла.');
    const records=[],warnings=[];
    for(let blockIndex=0;blockIndex<blocks.length;blockIndex++){
      const block=blocks[blockIndex],end=blockIndex+1<blocks.length?blocks[blockIndex+1].row:rows.length;
      let header=-1;
      for(let row=block.row+1;row<Math.min(end,block.row+9);row++)if(text(rows[row]?.[0])==='#'){header=row;break}
      if(header<0){warnings.push('Класс '+block.name+': не найдена строка с днями недели.');continue}
      const lessonRows=[];
      for(let row=header+1;row<end;row++){
        const value=Number(text(rows[row]?.[0]));
        if(Number.isInteger(value)&&value>=1&&value<=20)lessonRows.push({row,lesson:value});
      }
      for(let li=0;li<lessonRows.length;li++){
        const current=lessonRows[li],next=li+1<lessonRows.length?lessonRows[li+1].row:end;
        for(let dayIndex=0;dayIndex<DAYS.length;dayIndex++){
          const subjectColumn=1+dayIndex*2,roomColumn=subjectColumn+1,groups=[],rooms=[];
          for(let row=current.row;row<next;row++){
            for(const value of tokens(rows[row]?.[subjectColumn])){
              if(isDash(value))continue;
              if(looksLikeTeacher(value,knownTeachers)&&groups.length&&!groups[groups.length-1].teacher)groups[groups.length-1].teacher=value;
              else groups.push({subject:value,teacher:'',subjectRow:row});
            }
            for(const value of tokens(rows[row]?.[roomColumn]))if(!isDash(value)&&!/^Каб\.?$/i.test(value))rooms.push({value:roomName(value),row});
          }
          if(!groups.length)continue;
          groups.forEach((group,index)=>{
            let room='';
            if(rooms.length===groups.length)room=rooms[index].value;
            else if(rooms.length){
              const exact=rooms.find(item=>item.row===group.subjectRow);
              room=(exact||rooms[index]||{}).value||'';
            }
            records.push({'День':DAYS[dayIndex],'Урок':current.lesson,'Класс':block.name,'Предмет':group.subject,'Учитель':group.teacher,'Кабинет':room,'Группа':groups.length>1?String(index+1):'','Статус':group.teacher&&room?'ok':'check'});
          });
        }
      }
    }
    const unique=[],seen=new Set();
    for(const record of records){
      const key=[record['День'],record['Урок'],record['Класс'],record['Предмет'],record['Учитель'],record['Кабинет'],record['Группа']].join('\u001f');
      if(!seen.has(key)){seen.add(key);unique.push(record)}
    }
    if(!unique.length)throw new Error('В файле не найдено ни одного занятия. Файл не был загружен.');
    return {data:unique,warnings,classes:[...new Set(unique.map(x=>x['Класс']))],teachers:[...new Set(unique.map(x=>x['Учитель']).filter(Boolean))]};
  }
  return {parse};
});
