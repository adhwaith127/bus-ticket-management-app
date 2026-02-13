"""
mdb_debug_view.py  (FIXED)

TEMPORARY DEBUG VIEW — confirms mdbtools is reading your file correctly.
Does NOT write to DB. Dumps all table data to JSON so you can inspect it.

Steps:
1. Add to urls.py:
       from .mdb_debug_view import MdbDebugView
       path('debug-mdb/', MdbDebugView.as_view(), name='debug-mdb'),

2. Call with same FormData as import-mdb/ (file + optional password)

3. Read the API response directly — returns all table names + sample rows
   Also written to /tmp/mdb_debug_output.json on the server

FIX NOTE:
  Older mdbtools versions do NOT support -p flag on mdb-tables or mdb-export.
  Password is passed via MDB_JET_PASSWORD environment variable instead.
  This works across all mdbtools versions.
"""

import os
import json
import tempfile
import subprocess
import csv
import io

from rest_framework.views import APIView
from rest_framework.response import Response
from rest_framework import status
from rest_framework.parsers import MultiPartParser, FormParser


class MdbDebugView(APIView):
    parser_classes = [MultiPartParser, FormParser]

    def post(self, request):
        mdb_file = request.FILES.get('mdb_file')
        password = request.data.get('password', None)

        if not mdb_file:
            return Response({'message': 'No file provided.'}, status=status.HTTP_400_BAD_REQUEST)

        tmp_path = None

        try:
            # Save uploaded file to /tmp so mdbtools can read it from disk
            with tempfile.NamedTemporaryFile(suffix='.mdb', delete=False) as tmp:
                for chunk in mdb_file.chunks():
                    tmp.write(chunk)
                tmp_path = tmp.name

            # ---- Step 1: List ALL tables in the MDB ----
            tables_in_file = MdbDebugReader.list_tables(tmp_path, password)

            # ---- Step 2: Read every table, collect sample rows ----
            all_data = {}
            read_errors = {}

            for table_name in tables_in_file:
                try:
                    rows = MdbDebugReader.read_table(tmp_path, table_name, password)
                    all_data[table_name] = {
                        'row_count':   len(rows),
                        'columns':     list(rows[0].keys()) if rows else [],
                        'all_rows':    rows,
                    }
                except Exception as e:
                    read_errors[table_name] = str(e)

            # ---- Step 3: Write full output to JSON file on server ----
            output_path = '/tmp/mdb_debug_output.json'
            debug_output = {
                'file_name':    mdb_file.name,
                'tables_found': tables_in_file,
                'table_count':  len(tables_in_file),
                'data':         all_data,
                'read_errors':  read_errors,
            }
            with open(output_path, 'w') as f:
                json.dump(debug_output, f, indent=2, default=str)

            return Response({
                'message':      f'Debug complete. Output also written to {output_path}',
                'tables_found': tables_in_file,
                'table_count':  len(tables_in_file),
                'data':         all_data,
                'read_errors':  read_errors,
            }, status=status.HTTP_200_OK)

        except Exception as e:
            return Response(
                {'message': f'Debug failed: {str(e)}'},
                status=status.HTTP_500_INTERNAL_SERVER_ERROR
            )
        finally:
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)


class MdbDebugReader:

    @staticmethod
    def list_tables(mdb_path, password=None):
        """
        Lists all table names inside the .mdb file.
        Uses MDB_JET_PASSWORD env variable for password — NOT -p flag.
        """
        cmd = ['mdb-tables', '-1', mdb_path]

        env = os.environ.copy()
        if password:
            env['MDB_JET_PASSWORD'] = password

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=15, env=env)

        if result.returncode != 0:
            raise Exception(f'mdb-tables failed: {result.stderr}')

        tables = [t.strip() for t in result.stdout.strip().split('\n') if t.strip()]
        return tables

    @staticmethod
    def read_table(mdb_path, table_name, password=None):
        """
        Reads one table using mdb-export, returns list of dicts.
        Uses MDB_JET_PASSWORD env variable for password — NOT -p flag.
        """
        cmd = ['mdb-export', mdb_path, table_name]

        env = os.environ.copy()
        if password:
            env['MDB_JET_PASSWORD'] = password

        result = subprocess.run(cmd, capture_output=True, text=True, timeout=30, env=env)

        if result.returncode != 0:
            raise Exception(f'mdb-export failed: {result.stderr}')

        if not result.stdout.strip():
            return []

        reader = csv.DictReader(io.StringIO(result.stdout))
        return list(reader)