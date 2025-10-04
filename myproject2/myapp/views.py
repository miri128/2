from django.shortcuts import render
import openpyxl
from .forms import UploadFileForm
from .models import ItemProcess
from django.db.models import Sum
from django.utils import timezone
import datetime
import codecs
from django.db import transaction
import logging
import pandas as pd
from django.views.generic import TemplateView


logger = logging.getLogger(__name__)

class TopView(TemplateView):
    template_name = "top.html"

def upload_file(request):
    message = ""
    upload_count = 0
    last_uploaded_at = None

    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
            file = request.FILES['file']
            print("アップロード内容:",file.name,file.content_type)
            if file.name.lower().endswith('.txt'):
                try:
                    lines = codecs.iterdecode(file,'utf-8')
                    for line in lines:
                        print("line:", line)
                except Exception as e:
                    print("読み込みエラー", e)

                try:
                    with transaction.atomic():
                        ItemProcess.objects.all().delete()
                    file = request.FILES['file']
                    filename = file.name.lower()
                    DEPARTMENTS = ['PA', 'PB', 'PD', 'RA', 'RB', 'RC']

                    if filename.endswith('.xlsx'):
                        try:
                            wb = openpyxl.load_workbook(file, data_only=True)
                        except Exception as e:
                            message = "Excelファイルの読み込みに失敗しました。"
                            logger.error(f"Excel load error: {e}")
                            return render(request, 'upload.html', {
                                'form': form,
                                'message': message,
                                'uploaded_count': 0,
                                'last_uploaded_at': last_uploaded_at,
                                'current_data': [],
                            })
                        ws = wb.active

                        for row in ws.iter_rows(min_row=2):
                            department = (row[1].value or '').strip()
                            if department not in DEPARTMENTS:
                                continue

                            item_name = row[4].value
                            process_code = row[6].value
                            process_name = row[7].value
                            total_qty = row[8].value or 0
                            good_qty = row[9].value or 0
                            board_total_qty = row[13].value or 0
                            board_good_qty = row[14].value or 0
                            created_at = row[21].value

                            if not isinstance(created_at, datetime.datetime):
                                try:
                                    created_at = datetime.datetime.strptime(str(created_at), '%Y-%m-%d %H:%M:%S')
                                except Exception:
                                    created_at = timezone.now()
                                
                            try:
                                total_qty = int(total_qty)
                            except Exception:
                                total_qty = 0
                            try:
                                good_qty = int(good_qty)
                            except Exception:
                                good_qty = 0
                            try:
                                board_total_qty = int(board_total_qty)
                            except Exception:
                                board_total_qty = 0
                            try:
                                board_good_qty = int(board_good_qty)
                            except Exception:
                                board_good_qty = 0 

                            ItemProcess.objects.create(
                                    department=department,
                                    item_name=item_name,
                                    process_code=process_code,
                                    process_name=process_name,
                                    total_qty=total_qty,
                                    good_qty=good_qty,
                                    board_total_qty=board_total_qty,
                                    board_good_qty=board_good_qty,
                                    created_at=created_at,
                            )
                            upload_count += 1

                    elif filename.endswith('.txt'):
                        lines = codecs.iterdecode(file, 'utf-8')
                        for i, line in enumerate(lines):
                            if i == 0:
                                continue
                            values = line.strip().split('\t')
                            if len(values) < 22:
                                continue
                            department = values[1].strip()
                            if department not in DEPARTMENTS:
                                continue

                            item_name = values[4]
                            process_code = values[6]
                            process_name = values[7]
                            total_qty = values[8] or 0
                            good_qty = values[9] or 0
                            board_total_qty = values[13] or 0
                            board_good_qty = values[14] or 0
                            created_at = values[21]

                            try:
                                total_qty = int(total_qty)
                            except Exception:
                                total_qty = 0
                            try:
                                good_qty = int(good_qty)
                            except Exception:
                                good_qty = 0
                            try:
                                board_total_qty = int(board_total_qty)
                            except Exception:
                                board_total_qty = 0
                            try:
                                board_good_qty = int(board_good_qty)
                            except Exception:
                                board_good_qty = 0
                            try:
                                created_at = datetime.datetime.strptime(created_at, '%Y-%m-%d %H:%M:%S')
                            except Exception:
                                created_at = timezone.now()

                            ItemProcess.objects.create(
                                department=department,
                                item_name=item_name,
                                process_code=process_code,
                                process_name=process_name,
                                total_qty=total_qty,
                                good_qty=good_qty,
                                board_total_qty=board_total_qty,
                                board_good_qty=board_good_qty,
                                created_at=created_at,
                            )
                            upload_count += 1

                    else:
                        message = "対応しているファイル形式は .xlsx, .txt です。"

                        logger.warning(f"Unsupported file format uploaded: {filename}")
                    last_uploaded_at = timezone.now()
                    message = f"アップロード完了。{upload_count} 件のデータを取込みました。"

                except Exception as e:
                    logger.error(f"Upload processing error: {e}")
                    message = "アップロード処理中にエラーが発生しました。"

        else:
            form = UploadFileForm()

    else:
        form = UploadFileForm()
        last_record = ItemProcess.objects.order_by('-created_at').first()
        if last_record:
            last_uploaded_at = last_record.created_at

    current_data = ItemProcess.objects.all().order_by('-created_at')

    return render(request, 'upload.html', {
        'form': form,
        'message': message,
        'uploaded_count': upload_count,
        'last_uploaded_at': last_uploaded_at,
        'current_data': current_data,
    })

def  summary_view(request):
    ps = ItemProcess.objects.values('item_name', 'process_name', 'board_total_qty')
    df = pd.DataFrame.from_records(ps)
    pivot_table = df.pivot_table(
        index='item_name', 
        columns='process_name',
        values='board_total_qty', 
        aggfunc='sum',
        fill_value=0
    )

    process_alias_map = {
        '外し／詰め':'外し',
        '受け入れ':'受入',
        '基板変形検査':'反り',
        '基板スナップ割り':'ﾌﾞﾚｰｸ',
        'ブレーク':'ﾌﾞﾚｰｸ',
        '（外し／詰め）':'外し',
        '超音波ドライ洗浄(表)':'ﾄﾞﾗｲ洗浄',
        'ＱＣ外観チェック':'QC',
        'ロット構成':'ﾛｯﾄ構成',
        '出荷目視検査':'目視',
        '画像ﾎﾟｲﾝﾄ検査（表）':'画像',
        'ＤＭＣ印字':'DMC',
        'ＤＭＣ画像検査(表裏)':'画像',
        '基板外周検査(表)':'外周',
        '基板外周検査(裏)':'外周',
        '３Ｄ基板画像検査':'3D',
        '出荷目視／混入検査':'目視',
        '基板出荷画検':'目視',
        '基板画検 (表裏必須)':'画像',
        '基板出荷画検寸法必須':'画寸',
        '基板列削除':'列削除',
        '処置ﾎﾟｲﾝﾄ検査':'P',
        '指示事項伝達':'指示書',
        '基板外周抜取(70倍)':'外周抜',
        'AU付着抜取り検査':'Au',
        'QCﾛｯﾄ1':'ﾛｯﾄ1',
    }

    ordered_process = [
        '外し／詰め',
        '受け入れ',
        '基板変形検査',
        '基板スナップ割り',
        'ブレーク',
        '（外し／詰め）',
        '超音波ドライ洗浄(表)',
        'ＱＣ外観チェック',
        'ロット構成',
        '出荷目視検査',
        '画像ﾎﾟｲﾝﾄ検査（表）',
        'ＤＭＣ印字',
        'ＤＭＣ画像検査(表裏)',
        '基板外周検査(表)',
        '基板外周検査(裏)',
        '３Ｄ基板画像検査',
        '出荷目視／混入検査',
        '基板出荷画検',
        '基板画検 (表裏必須)',
        '基板出荷画検寸法必須',
        '基板列削除',
        '処置ﾎﾟｲﾝﾄ検査',
        '指示事項伝達',
        '基板外周抜取(70倍)',
        'AU付着抜取り検査',
        'QCﾛｯﾄ1',
    ]

    nonzero_columns = [col for col in ordered_process if col in pivot_table.columns]
    pivot_table = pivot_table.reindex(columns=nonzero_columns, fill_value=0)
    pivot_data = pivot_table.reset_index().to_dict(orient='records')
    column_headers = ['item_name'] + nonzero_columns
    display_headers = ['品名'] + [process_alias_map.get(col, col) for col in nonzero_columns]
 
    return render(request, 'summary.html', {
        'pivot_data': pivot_data,
        'column_headers': column_headers,
        'display_headers': display_headers
    })
       