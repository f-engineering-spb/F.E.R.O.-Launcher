"""HTML-оболочка листа Excel: ПОБАЙТОВЫЙ ПОРТ из server.py.

JS-протокол iframe (launcher-sheet-zoom/rotate/fit/hand + ответ
launcher-sheet-fitted) — контракт с app/frontend/app.js.
Не менять разметку/скрипт без обновления фронта.
"""

from __future__ import annotations


def render_sheet_page(*, title: str, cols: str, rows: list,
                      table_width: int, notice: str) -> str:
    """Собрать полную страницу листа. rows — список строк таблицы."""
    import html
    _sheet = type("Sheet", (), {})()
    _sheet.title = title
    sheet = _sheet
    page = f'''<!doctype html><html lang="ru"><head><meta charset="utf-8"><title>{html.escape(sheet.title)}</title>
<style>
html,body{{margin:0;min-width:max-content;background:#fff;color:#20262d;font:11px "Segoe UI",Arial,sans-serif;overflow:auto}}
#sheet-canvas{{position:relative;transform-origin:0 0}}#sheet{{position:absolute;left:0;top:0;transform-origin:50% 50%}}
.notice{{position:sticky;top:0;z-index:2;padding:6px 10px;border-bottom:1px solid #d3dde4;background:#f7fafc;color:#607080;font-size:11px}}
table{{border-collapse:collapse;table-layout:fixed;width:{table_width}px}}col.row-number{{width:38px}}
th,td{{box-sizing:border-box;border:1px solid #cbd5dc;padding:2px 4px;vertical-align:top;white-space:pre-wrap;overflow-wrap:break-word}}
th{{position:sticky;left:0;z-index:1;background:#f1f5f7;color:#657687;font:10px "Segoe UI",Arial,sans-serif;text-align:right}}td{{overflow:hidden}}
</style></head><body><div id="sheet-canvas"><div id="sheet"><table><colgroup>{cols}</colgroup><tbody>{''.join(rows)}</tbody></table></div></div>
<script>const canvas=document.getElementById('sheet-canvas'),sheet=document.getElementById('sheet');let width=0,height=0,padding=0,scale=1,rotation=0,hand=true,dragging=false,startX=0,startY=0,startLeft=0,startTop=0;function dimensions(){{return Math.abs(rotation%180)===90?{{width:height,height:width}}:{{width,height}}}}function zoom(value){{if(!width){{width=sheet.offsetWidth;height=sheet.offsetHeight}}scale=Math.max(.35,Math.min(3,value));padding=Math.max(innerWidth,innerHeight);const size=dimensions();canvas.style.width=(size.width*scale+padding*2)+'px';canvas.style.height=(size.height*scale+padding*2)+'px';sheet.style.transform='translate('+(padding+size.width*scale/2-width/2)+'px,'+(padding+size.height*scale/2-height/2)+'px) rotate('+rotation+'deg) scale('+scale+')'}}function fit(){{if(!width)zoom(1);const size=dimensions(),value=Math.min(1,(innerWidth-48)/size.width,(innerHeight-48)/size.height);zoom(value);requestAnimationFrame(()=>{{scrollTo(padding,padding);parent.postMessage({{type:'launcher-sheet-fitted',value:scale}},'*')}})}}function cursor(){{document.body.style.cursor=hand?(dragging?'grabbing':'grab'):'default'}}addEventListener('load',()=>{{zoom(1);cursor();requestAnimationFrame(()=>scrollTo(padding,padding))}});addEventListener('wheel',event=>{{if(!event.ctrlKey)return;event.preventDefault();zoom(scale*(event.deltaY<0?1.12:.89))}},{{passive:false}});addEventListener('pointerdown',event=>{{if(!hand||event.button!==0)return;dragging=true;startX=event.clientX;startY=event.clientY;startLeft=scrollX;startTop=scrollY;document.body.setPointerCapture?.(event.pointerId);cursor();event.preventDefault()}});addEventListener('pointermove',event=>{{if(!dragging)return;scrollTo(startLeft-(event.clientX-startX),startTop-(event.clientY-startY))}});addEventListener('pointerup',event=>{{if(!dragging)return;dragging=false;document.body.releasePointerCapture?.(event.pointerId);cursor()}});addEventListener('pointercancel',()=>{{dragging=false;cursor()}});addEventListener('keydown',event=>{{if(event.key==='Escape')parent.postMessage({{type:'launcher-escape'}},'*')}});addEventListener('dblclick',()=>{{parent.postMessage({{type:'launcher-toggle-full-view'}},'*')}});addEventListener('message',event=>{{if(!event.data)return;if(event.data.type==='launcher-sheet-zoom')zoom(event.data.value);if(event.data.type==='launcher-sheet-fit')fit();if(event.data.type==='launcher-sheet-rotate'){{rotation=((Number(event.data.value)||0)%360+360)%360;zoom(scale)}}if(event.data.type==='launcher-sheet-hand'){{hand=Boolean(event.data.value);dragging=false;cursor()}}}});</script></body></html>'''
    return page
