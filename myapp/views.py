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


logger = logging.getLogger(__name__)

def upload_file(request):
    message = ""
    upload_count = 0
    last_uploaded_at = None

    if request.method == 'POST':
        form = UploadFileForm(request.POST, request.FILES)
        if form.is_valid():
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

                    elif filename.endswith('.txt') or filename.endswith('.csv'):
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
                        message = "対応しているファイル形式は .xlsx, .csv, .txt です。"

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
    pivot_table = df.pivot_table(index='item_name', columns='process_name', values='board_total_qty', aggfunc='sum', fill_value=0)
    pivot_data = pivot_table.reset_index().to_dict(orient='records')
    column_headers = ['item_name'] + list(pivot_table.columns)
    display_headers = ['品名'] + list(pivot_table.columns)

    return render(request, 'summary.html', {
        'pivot_data': pivot_data,
        'column_headers': column_headers,
        'display_headers': display_headers
    })
       