#!/usr/bin/env python3
"""
Simple HTTP server for the Manufacturing Analytics Dashboard
Serves the dashboard and JSON data
"""

import http.server
import socketserver
import os
import sys
sys.stdout.reconfigure(encoding='utf-8')
from pathlib import Path
import json
from openpyxl import load_workbook

PORT = 8000

class MyHTTPRequestHandler(http.server.SimpleHTTPRequestHandler):

    def end_headers(self):
        self.send_header('Cache-Control', 'no-store, no-cache, must-revalidate, max-age=0')
        self.send_header('Pragma', 'no-cache')
        self.send_header('Expires', '0')
        self.send_header('Access-Control-Allow-Origin', '*')
        super().end_headers()

    def do_OPTIONS(self):
        self.send_response(200)
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.end_headers()

    def do_GET(self):
        if self.path.startswith('/api/get-car-entries'):
            self.handle_get_car_entries()
        elif self.path.startswith('/api/get-query-entries'):
            self.handle_get_query_entries()
        elif self.path.startswith('/api/get-maintenance-entries'):
            self.handle_get_maintenance_entries()
        elif self.path.startswith('/api/get-keshkomi-entries'):
            self.handle_get_keshkomi_entries()
        elif self.path.startswith('/api/get-punch-entries'):
            self.handle_get_punch_entries()
        else:
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/update-safety':
            self.handle_update_safety()
        elif self.path == '/api/update-yearly-target':
            self.handle_update_yearly_target()
        elif self.path == '/api/add-car-entry':
            self.handle_add_car_entry()
        elif self.path == '/api/update-car-actual':
            self.handle_update_car_actual()
        elif self.path == '/api/add-query-entry':
            self.handle_add_query_entry()
        elif self.path == '/api/update-query-stage':
            self.handle_update_query_stage()
        elif self.path == '/api/update-query-summary':
            self.handle_update_query_summary()
        elif self.path == '/api/update-query-milestone':
            self.handle_update_query_milestone()
        elif self.path == '/api/add-maintenance-entry':
            self.handle_add_maintenance_entry()
        elif self.path == '/api/update-maintenance-entry':
            self.handle_update_maintenance_entry()
        elif self.path == '/api/add-keshkomi-entry':
            self.handle_add_keshkomi_entry()
        elif self.path == '/api/update-keshkomi-status':
            self.handle_update_keshkomi_status()
        elif self.path == '/api/delete-keshkomi-entry':
            self.handle_delete_keshkomi_entry()
        else:
            self.send_response(404)
            self.end_headers()

    # ── SAFETY ──────────────────────────────────────────────────────
    def handle_update_safety(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            month  = data.get('month')
            status = data.get('status')
            days   = data.get('days')
            lwd_date = data.get('lwd_date', '')
            type_of_month = data.get('type_of_month', '')

            print(f"📝 Safety update received: month={month}, status={status}, days={days}")

            wb = load_workbook('data/Safety_Status.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                cell_val = row[0].value
                if cell_val and str(cell_val).strip().lower() == month.lower():
                    print(f"✅ Match found for {month}, writing values...")
                    if status == 'red':
                        row[1].value = days if days else 1
                        row[2].value = days if days else 1
                    elif status == 'blue':
                        row[1].value = 1
                        row[2].value = 0
                    else:
                        row[1].value = 0
                        row[2].value = 0
                    break

            wb.save('data/Safety_Status.xlsx')
            print(f"💾 Safety Excel saved successfully")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error in update-safety: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── YEARLY TARGET ────────────────────────────────────────────────
    def handle_update_yearly_target(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)
            new_target = float(data.get('yearly_target', 0))

            wb = load_workbook('data/IHTM_Revenue.xlsx')
            ws = wb.active

            headers = [cell.value for cell in ws[1]]
            target_col = next((i + 1 for i, h in enumerate(headers) if h and 'target' in str(h).lower()), None)

            if target_col:
                monthly = new_target / 12
                for row in ws.iter_rows(min_row=2):
                    if row[0].value:
                        row[target_col - 1].value = round(monthly, 2)

            wb.save('data/IHTM_Revenue.xlsx')
            print(f"💾 Yearly target updated to {new_target} M")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating yearly target: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── CAR USAGE — ADD ─────────────────────────────────────────────
    def handle_add_car_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            wb = load_workbook('data/Car_Usage.xlsx')
            ws = wb.active

            next_row = ws.max_row + 1

            ws.cell(next_row, 1).value = int(data.get('sl_no')) if data.get('sl_no') else next_row - 1
            ws.cell(next_row, 2).value = data.get('date')
            ws.cell(next_row, 3).value = data.get('out_time')
            ws.cell(next_row, 4).value = data.get('tentative')
            ws.cell(next_row, 5).value = data.get('actual')
            ws.cell(next_row, 6).value = data.get('project')
            ws.cell(next_row, 7).value = data.get('name')
            ws.cell(next_row, 8).value = data.get('tm_no')

            wb.save('data/Car_Usage.xlsx')
            print(f"✅ Car entry saved: {data.get('name')} on {data.get('date')}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error saving car entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── CAR USAGE — UPDATE ACTUAL RETURN TIME ────────────────────────
    def handle_update_car_actual(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            sl_no = data.get('sl_no')
            actual_return = data.get('actual_return')

            wb = load_workbook('data/Car_Usage.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                if row[0].value and str(row[0].value) == str(sl_no):
                    row[4].value = actual_return
                    break

            wb.save('data/Car_Usage.xlsx')
            print(f"✅ Car actual return updated: Sl No {sl_no}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating car entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── CAR USAGE — GET ─────────────────────────────────────────────
    def handle_get_car_entries(self):
        try:
            wb = load_workbook('data/Car_Usage.xlsx')
            ws = wb.active
            entries = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    entries.append({
                        'sl_no':      row[0],
                        'date':       str(row[1]) if row[1] else '',
                        'out_time':   row[2] or '',
                        'tentative':  row[3] or '',
                        'actual':     row[4] or '',
                        'project':    row[5] or '',
                        'name':       row[6] or '',
                        'tm_no':      row[7] or '',
                    })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': entries}).encode())

        except Exception as e:
            print(f"❌ Error reading car entries: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': [], 'error': str(e)}).encode())

    # ── QUERY — ADD ─────────────────────────────────────────────────
    def handle_add_query_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            wb = load_workbook('data/Query_Analytics.xlsx')
            ws = wb.active

            next_row = ws.max_row + 1

            # Column mapping for Query_Analytics:
            # 1: Plant, 2: Stage, 3: Project Name, 4: Summary, 5: Current Stage Date
            # Extended: 6: Customer Name, 7: Leader, 8: Tool No, 9: Date Query, 10: Date Spec, 11: Date Concept, 12: Date Estim, 13: Date PO

            # Ensure columns exist
            max_col = ws.max_column
            headers = [cell.value for cell in ws[1]] if ws.max_row >= 1 else []

            # Write to existing columns or extend
            ws.cell(next_row, 1).value = data.get('plant', '')
            ws.cell(next_row, 2).value = data.get('current_stage', 'Query')  # Stage
            ws.cell(next_row, 3).value = data.get('project_name', '')
            ws.cell(next_row, 4).value = data.get('summary', '')
            ws.cell(next_row, 5).value = data.get('current_stage_date', '')

            # Extended columns
            if max_col >= 6: ws.cell(next_row, 6).value = data.get('customer_name', '')
            if max_col >= 7: ws.cell(next_row, 7).value = data.get('leader', '')
            if max_col >= 8: ws.cell(next_row, 8).value = data.get('tool_no', '')
            if max_col >= 9: ws.cell(next_row, 9).value = data.get('date_query', '')
            if max_col >= 10: ws.cell(next_row, 10).value = data.get('date_spec', '')
            if max_col >= 11: ws.cell(next_row, 11).value = data.get('date_concept', '')
            if max_col >= 12: ws.cell(next_row, 12).value = data.get('date_estim', '')
            if max_col >= 13: ws.cell(next_row, 13).value = data.get('date_po', '')

            wb.save('data/Query_Analytics.xlsx')
            print(f"✅ Query entry saved: {data.get('project_name')}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error saving query entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── QUERY — UPDATE STAGE (includes date) ───────────────────────
    def handle_update_query_stage(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            project_name = data.get('project_name')
            new_stage = data.get('current_stage')
            stage_date = data.get('stage_date')

            wb = load_workbook('data/Query_Analytics.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                if row[2].value and str(row[2].value).strip() == str(project_name).strip():
                    # Update Stage (col 2)
                    row[1].value = new_stage
                    # Update Current Stage Date (col 5)
                    if stage_date:
                        row[4].value = stage_date
                    # Also update the milestone date column based on stage
                    stage_map = {
                        'Query': 9, 'Speculation': 10, 'Concept': 11,
                        'Estimation': 12, 'PO/Budget': 13
                    }
                    if new_stage in stage_map and stage_date:
                        col_idx = stage_map[new_stage]
                        if ws.max_column >= col_idx:
                            ws.cell(row[0].row, col_idx).value = stage_date
                    break

            wb.save('data/Query_Analytics.xlsx')
            print(f"✅ Query stage updated: {project_name} → {new_stage} on {stage_date}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating query stage: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── QUERY — GET ──────────────────────────────────────────────────
    def handle_get_query_entries(self):
        try:
            wb = load_workbook('data/Query_Analytics.xlsx')
            ws = wb.active
            entries = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    entries.append({
                        'plant':              row[0] or '',
                        'stage':              row[1] or '',
                        'project_name':       row[2] or '',
                        'summary':            row[3] or '',
                        'current_stage_date': str(row[4]) if row[4] else '',
                        'customer_name':      row[5] or '' if len(row) > 5 else '',
                        'leader':             row[6] or '' if len(row) > 6 else '',
                        'tool_no':            row[7] or '' if len(row) > 7 else '',
                        'date_query':         str(row[8]) if len(row) > 8 and row[8] else '',
                        'date_spec':          str(row[9]) if len(row) > 9 and row[9] else '',
                        'date_concept':       str(row[10]) if len(row) > 10 and row[10] else '',
                        'date_estim':         str(row[11]) if len(row) > 11 and row[11] else '',
                        'date_po':            str(row[12]) if len(row) > 12 and row[12] else '',
                    })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': entries}).encode())

        except Exception as e:
            print(f"❌ Error reading query entries: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': [], 'error': str(e)}).encode())

    # ── QUERY — UPDATE SUMMARY ───────────────────────────────────────
    def handle_update_query_summary(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            project_name = data.get('project_name')
            summary = data.get('summary')

            wb = load_workbook('data/Query_Analytics.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                if row[2].value and str(row[2].value).strip() == str(project_name).strip():
                    row[3].value = summary
                    break

            wb.save('data/Query_Analytics.xlsx')
            print(f"✅ Query summary updated: {project_name}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating query summary: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── QUERY — UPDATE MILESTONE ──────────────────────────────────────
    def handle_update_query_milestone(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            project_name = data.get('project_name')
            milestone = data.get('milestone')
            date = data.get('date')

            wb = load_workbook('data/Query_Analytics.xlsx')
            ws = wb.active

            # Map milestone names to column indices
            milestone_map = {
                'Query': 9, 'Speculation': 10, 'Concept': 11,
                'Estimation': 12, 'PO/Budget': 13
            }
            col_idx = milestone_map.get(milestone, 6)

            for row in ws.iter_rows(min_row=2):
                if row[2].value and str(row[2].value).strip() == str(project_name).strip():
                    if ws.max_column >= col_idx:
                        ws.cell(row[0].row, col_idx).value = date
                    break

            wb.save('data/Query_Analytics.xlsx')
            print(f"✅ Query milestone updated: {project_name} → {milestone} = {date}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating query milestone: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── MAINTENANCE — ADD ────────────────────────────────────────────
    def handle_add_maintenance_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            wb = load_workbook('data/Car_Maintainace_Schedule.xlsx')
            ws = wb.active

            next_row = ws.max_row + 1

            ws.cell(next_row, 1).value = data.get('name')
            ws.cell(next_row, 2).value = data.get('cleanDate')
            ws.cell(next_row, 3).value = 'Yes' if data.get('cleanAck') else 'No'
            ws.cell(next_row, 4).value = data.get('confDate')
            ws.cell(next_row, 5).value = 'Yes' if data.get('confAck') else 'No'
            # Add week column if exists
            if ws.max_column >= 6:
                ws.cell(next_row, 6).value = data.get('week', '')

            wb.save('data/Car_Maintainace_Schedule.xlsx')
            print(f"✅ Maintenance entry saved: {data.get('name')}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error saving maintenance entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── MAINTENANCE — UPDATE ─────────────────────────────────────────
    def handle_update_maintenance_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            name = data.get('name')
            field = data.get('field')
            value = data.get('value')

            wb = load_workbook('data/Car_Maintainace_Schedule.xlsx')
            ws = wb.active

            col_map = {
                'cleanDate': 2,
                'cleanAck': 3,
                'confDate': 4,
                'confAck': 5,
            }

            col_idx = col_map.get(field)
            if not col_idx:
                raise ValueError(f"Unknown field: {field}")

            for row in ws.iter_rows(min_row=2):
                if row[0].value and str(row[0].value).strip() == str(name).strip():
                    if field in ['cleanAck', 'confAck']:
                        row[col_idx - 1].value = 'Yes' if value else 'No'
                    else:
                        row[col_idx - 1].value = value
                    break

            wb.save('data/Car_Maintainace_Schedule.xlsx')
            print(f"✅ Maintenance entry updated: {name}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating maintenance entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── MAINTENANCE — GET ────────────────────────────────────────────
    def handle_get_maintenance_entries(self):
        try:
            wb = load_workbook('data/Car_Maintainace_Schedule.xlsx')
            ws = wb.active
            entries = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    entries.append({
                        'name':     row[0] or '',
                        'cleanDate': str(row[1]) if row[1] else '',
                        'cleanAck': (str(row[2] or '').lower() == 'yes'),
                        'confDate':  str(row[3]) if row[3] else '',
                        'confAck':   (str(row[4] or '').lower() == 'yes'),
                        'week':      row[5] or '' if len(row) > 5 else '',
                    })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': entries}).encode())

        except Exception as e:
            print(f"❌ Error reading maintenance entries: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': [], 'error': str(e)}).encode())

    # ── KESHKOMI — ADD ──────────────────────────────────────────────
    def handle_add_keshkomi_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            wb = load_workbook('data/KeshKomi.xlsx')
            ws = wb.active

            next_row = ws.max_row + 1

            ws.cell(next_row, 1).value = int(data.get('sl_no')) if data.get('sl_no') else next_row - 1
            ws.cell(next_row, 2).value = data.get('kadai_point')
            ws.cell(next_row, 3).value = data.get('action_plan')
            ws.cell(next_row, 4).value = data.get('responsibility')
            ws.cell(next_row, 5).value = data.get('target_date')
            ws.cell(next_row, 6).value = data.get('status', 'Incomplete')
            ws.cell(next_row, 7).value = bool(data.get('high_priority', False))

            wb.save('data/KeshKomi.xlsx')
            print(f"✅ KeshKomi entry saved: {data.get('kadai_point')}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error saving KeshKomi entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── KESHKOMI — GET ──────────────────────────────────────────────
    def handle_get_keshkomi_entries(self):
        try:
            wb = load_workbook('data/KeshKomi.xlsx')
            ws = wb.active
            entries = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    entries.append({
                        'sl_no':          row[0],
                        'kadai_point':    row[1] or '',
                        'action_plan':    row[2] or '',
                        'responsibility': row[3] or '',
                        'target_date':    str(row[4]) if row[4] else '',
                        'status':         row[5] or 'Incomplete',
                        'high_priority':  bool(row[6]) if len(row) > 6 and row[6] is not None else False,
                    })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': entries}).encode())

        except Exception as e:
            print(f"❌ Error reading KeshKomi entries: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': [], 'error': str(e)}).encode())

    # ── KESHKOMI — UPDATE STATUS (includes high_priority) ───────────
    def handle_update_keshkomi_status(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            sl_no = data.get('sl_no')
            new_status = data.get('status')
            high_priority = data.get('high_priority')

            wb = load_workbook('data/KeshKomi.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                if row[0].value and str(row[0].value) == str(sl_no):
                    row[5].value = new_status
                    if high_priority is not None and len(row) > 6:
                        row[6].value = bool(high_priority)
                    break

            wb.save('data/KeshKomi.xlsx')
            print(f"✅ KeshKomi status updated: Sl No {sl_no} → {new_status}, HP={high_priority}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error updating KeshKomi status: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── KESHKOMI — DELETE ────────────────────────────────────────────
    def handle_delete_keshkomi_entry(self):
        try:
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            data = json.loads(body)

            sl_no = data.get('sl_no')

            wb = load_workbook('data/KeshKomi.xlsx')
            ws = wb.active

            for row in ws.iter_rows(min_row=2):
                if row[0].value and str(row[0].value) == str(sl_no):
                    ws.delete_rows(row[0].row, 1)
                    break

            wb.save('data/KeshKomi.xlsx')
            print(f"✅ KeshKomi entry deleted: Sl No {sl_no}")

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': True}).encode())

        except Exception as e:
            print(f"❌ Error deleting KeshKomi entry: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'success': False, 'error': str(e)}).encode())

    # ── PUNCH POINTS — GET ──────────────────────────────────────────
    def handle_get_punch_entries(self):
        try:
            wb = load_workbook('data/Project_Punch_Points.xlsx')
            ws = wb.active
            entries = []
            for row in ws.iter_rows(min_row=2, values_only=True):
                if any(cell is not None for cell in row):
                    entries.append({
                        'item': row[0] or '',
                        'month': row[1] or '',
                        'punch_points': row[2] or 0,
                    })

            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': entries}).encode())

        except Exception as e:
            print(f"❌ Error reading punch entries: {e}")
            self.send_response(500)
            self.send_header('Content-Type', 'application/json')
            self.end_headers()
            self.wfile.write(json.dumps({'entries': [], 'error': str(e)}).encode())

    def log_message(self, format, *args):
        print(f'[{self.log_date_time_string()}] {format % args}')


def main():
    script_dir = Path(__file__).resolve().parent
    os.chdir(script_dir)

    print(f"""
╔══════════════════════════════════════════════════════════╗
║   IHTM Analytics Dashboard - Development Server         ║
╚══════════════════════════════════════════════════════════╝

📊 Server starting...
🌐 URL: http://localhost:{PORT}
📁 Directory: {script_dir}

Files being served:
  ✓ dashboard.html       (main dashboard)
  ✓ analytics_data.json  (data)
  ✓ data/*.xlsx          (Excel files)

API Endpoints:
  ✓ POST /api/update-safety            (safety calendar → Excel)
  ✓ POST /api/update-yearly-target     (yearly target → Excel)
  ✓ POST /api/add-car-entry            (car usage form → Excel)
  ✓ POST /api/update-car-actual        (update actual return time → Excel)
  ✓ GET  /api/get-car-entries          (read car usage from Excel)
  ✓ POST /api/add-query-entry          (query monitoring form → Excel)
  ✓ GET  /api/get-query-entries        (read query details from Excel)
  ✓ POST /api/update-query-stage       (update query stage + date → Excel)
  ✓ POST /api/update-query-summary     (update query summary → Excel)
  ✓ POST /api/update-query-milestone   (update query milestone → Excel)
  ✓ POST /api/add-maintenance-entry    (car maintenance form → Excel)
  ✓ POST /api/update-maintenance-entry (update maintenance → Excel)
  ✓ GET  /api/get-maintenance-entries  (read maintenance from Excel)
  ✓ POST /api/add-keshkomi-entry       (KeshKomi form → Excel)
  ✓ GET  /api/get-keshkomi-entries     (read KeshKomi from Excel)
  ✓ POST /api/update-keshkomi-status   (update KeshKomi status + HP → Excel)
  ✓ POST /api/delete-keshkomi-entry    (delete KeshKomi entry → Excel)
  ✓ GET  /api/get-punch-entries        (read punch points from Excel)

⚠️  Press CTRL+C to stop the server

""")

    with socketserver.TCPServer(("", PORT), MyHTTPRequestHandler) as httpd:
        httpd.allow_reuse_address = True
        try:
            httpd.serve_forever()
        except KeyboardInterrupt:
            print("\n🛑 Server stopped.")
            httpd.shutdown()

if __name__ == '__main__':
    main()