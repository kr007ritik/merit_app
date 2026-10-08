from datetime import datetime
import os
from flask import Flask, flash, redirect, render_template, request, send_file
import pandas as pd

app = Flask(__name__)
app.secret_key = 'merit_portal_secret_key'

UPLOAD_FOLDER = os.path.join(os.path.dirname(__file__), 'uploads')
os.makedirs(UPLOAD_FOLDER, exist_ok=True)


def process_merit_logic(file_path, output_path):
  # Read Excel File
  xls = pd.ExcelFile(file_path)
  sheet_names = [s.strip() for s in xls.sheet_names]

  if 'raw_data' not in sheet_names or 'vacency' not in sheet_names:
    raise ValueError(
        f"Excel file me `raw_data` aur `vacency` sheets honi chahiye. Found:"
        f' {sheet_names}'
    )

  df_raw = pd.read_excel(xls, sheet_name='raw_data')
  df_vac = pd.read_excel(xls, sheet_name='vacency')

  # Clean Column Names (strip spaces)
  df_raw.columns = df_raw.columns.str.strip()
  df_vac.columns = df_vac.columns.str.strip()

  # Mandatory Columns Verification
  required_raw_cols = [
      'student_name',
      'mobile_no',
      'dob_(dd-mm-yyyy)',
      '10th %',
      '12th %',
      'unique_id',
      'subject',
      'college_preferance_1',
  ]
  missing_cols = [col for col in required_raw_cols if col not in df_raw.columns]
  if missing_cols:
    raise ValueError(f"`raw_data` sheet me missing columns: {missing_cols}")

  # Clean Data
  df_raw['student_name'] = df_raw['student_name'].astype(str).str.strip()
  df_raw['mobile_no'] = df_raw['mobile_no'].astype(str).str.strip()
  df_raw['dob_str'] = df_raw['dob_(dd-mm-yyyy)'].astype(str).str.strip()

  df_raw['12th %'] = pd.to_numeric(df_raw['12th %'], errors='coerce').fillna(0)
  df_raw['10th %'] = pd.to_numeric(df_raw['10th %'], errors='coerce').fillna(0)

  # Date Parsing
  df_raw['dob_dt'] = pd.to_datetime(
      df_raw['dob_str'], format='%d-%m-%Y', errors='coerce'
  )

  # Student Key
  df_raw['student_key'] = (
      df_raw['student_name'] + '_' + df_raw['mobile_no'] + '_' + df_raw['dob_str']
  )

  # Vacancy Dictionary
  df_vac['college_name_clean'] = df_vac['college_name'].astype(str).str.strip()
  df_vac['subject_clean'] = df_vac['subject'].astype(str).str.strip()

  current_vac = {}
  for _, row in df_vac.iterrows():
    key = (row['college_name_clean'], row['subject_clean'])
    current_vac[key] = int(row['vacent_seat'])

  # Merit Ranking
  student_merit_list = df_raw.groupby('student_key').first().reset_index()
  student_merit_list = student_merit_list.sort_values(
      by=['12th %', '10th %', 'dob_dt', 'student_name'],
      ascending=[False, False, True, True],
  ).reset_index(drop=True)

  student_merit_list['overall_merit_rank'] = student_merit_list.index + 1

  # Allotment Logic
  allotment_results = []
  for idx, st in student_merit_list.iterrows():
    st_key = st['student_key']
    rank = st['overall_merit_rank']

    st_apps = df_raw[df_raw['student_key'] == st_key].copy()
    selected = False
    allotted_info = {}

    for _, app in st_apps.iterrows():
      if selected:
        break

      sub = str(app['subject']).strip()
      form_no = app['unique_id']

      for pref_num in [1, 2, 3]:
        pref_col = f'college_preferance_{pref_num}'
        if pref_col not in app:
          continue

        col_name = app[pref_col]
        if (
            pd.isna(col_name)
            or str(col_name).strip() == ''
            or str(col_name).strip().lower() == 'nan'
        ):
          continue

        col_name_clean = str(col_name).strip()
        vac_key = (col_name_clean, sub)

        if current_vac.get(vac_key, 0) > 0:
          selected = True
          current_vac[vac_key] -= 1

          allotted_info = {
              'student_key': st_key,
              'overall_merit_rank': rank,
              'student_name': app['student_name'],
              'mobile_no': app['mobile_no'],
              'dob_(dd-mm-yyyy)': app['dob_str'],
              '10th %': app['10th %'],
              '12th %': app['12th %'],
              'allocated_form_no': form_no,
              'allocated_subject': sub,
              'allocated_college': col_name_clean,
              'preference_matched': f'Preference {pref_num}',
              'selection_status': 'Selected',
          }
          break

    if not selected:
      allotted_info = {
          'student_key': st_key,
          'overall_merit_rank': rank,
          'student_name': st['student_name'],
          'mobile_no': st['mobile_no'],
          'dob_(dd-mm-yyyy)': st['dob_str'],
          '10th %': st['10th %'],
          '12th %': st['12th %'],
          'allocated_form_no': 'N/A',
          'allocated_subject': 'N/A',
          'allocated_college': 'N/A',
          'preference_matched': 'N/A',
          'selection_status': 'Not Selected',
      }

    allotment_results.append(allotted_info)

  df_merit_summary = pd.DataFrame(allotment_results)

  # Cutoff Sheet
  selected_students = df_merit_summary[
      df_merit_summary['selection_status'] == 'Selected'
  ].copy()
  cutoff_df = (
      selected_students.groupby(['allocated_college', 'allocated_subject'])
      .agg(
          total_selected=('student_key', 'count'),
          max_selected_12th_pct=('12th %', 'max'),
          cutoff_12th_pct=('12th %', 'min'),
          last_selected_10th_pct=('10th %', 'min'),
      )
      .reset_index()
  )

  cutoff_df.rename(
      columns={
          'allocated_college': 'college_name',
          'allocated_subject': 'subject',
          'max_selected_12th_pct': 'max_selected_12th_%',
          'cutoff_12th_pct': 'cutoff_12th_%',
          'last_selected_10th_pct': 'last_selected_10th_%',
      },
      inplace=True,
  )

  # Vacancy Sheet Update
  df_updated_vac = df_vac.copy()

  def get_remaining_seats(row):
    key = (str(row['college_name']).strip(), str(row['subject']).strip())
    return current_vac.get(key, 0)

  df_updated_vac['remaining_vacent_seat'] = df_updated_vac.apply(
      get_remaining_seats, axis=1
  )
  df_updated_vac['seats_filled'] = (
      df_updated_vac['vacent_seat'] - df_updated_vac['remaining_vacent_seat']
  )
  df_updated_vac.drop(
      columns=['college_name_clean', 'subject_clean'],
      inplace=True,
      errors='ignore',
  )

  # Save Excel
  with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    df_merit_summary.to_excel(writer, sheet_name='merit_list', index=False)
    cutoff_df.to_excel(writer, sheet_name='cutoff_details', index=False)
    df_raw.to_excel(writer, sheet_name='raw_data', index=False)
    df_updated_vac.to_excel(writer, sheet_name='updated_vacency', index=False)


@app.route('/', methods=['GET', 'POST'])
def index():
  if request.method == 'POST':
    if 'excel_file' not in request.files:
      flash('Kripya Excel file select karein!')
      return redirect(request.url)

    file = request.files['excel_file']
    if file.filename == '':
      flash('Koi file select nahi hui!')
      return redirect(request.url)

    if file and (
        file.filename.endswith('.xlsx') or file.filename.endswith('.xls')
    ):
      input_filepath = os.path.join(UPLOAD_FOLDER, 'uploaded_data.xlsx')
      output_filepath = os.path.join(UPLOAD_FOLDER, 'merit_list_result.xlsx')

      file.save(input_filepath)

      try:
        process_merit_logic(input_filepath, output_filepath)
        return send_file(
            output_filepath,
            as_attachment=True,
            download_name='Generated_Merit_List.xlsx',
        )
      except Exception as e:
        flash(f'Processing Error: {str(e)}')
        return redirect(request.url)
    else:
      flash('Keval .xlsx ya .xls Excel file upload karein!')
      return redirect(request.url)

  return render_template('index.html')


if __name__ == '__main__':
  port = int(os.environ.get('PORT', 5000))
  app.run(host='0.0.0.0', port=port)