import os
import uuid
from flask import Flask, request, jsonify, render_template_string, send_file
import fitz  # PyMuPDF
import pandas as pd
from fpdf import FPDF
import pdfplumber
from pdf2docx import Converter
import docx
import zipfile
import shutil
import openpyxl
from openpyxl.styles import Border, Side, Alignment

# --- NEW IMPORTS FOR WORD TO PDF ---
import pythoncom
from docx2pdf import convert as docx2pdf_convert

# --- NEW IMPORTS FOR OCR ---
import pytesseract
from PIL import Image
import time # NEW: Needed for user tracking

# IMPORTANT FOR WINDOWS USERS: 
# You must install the Tesseract executable and point pytesseract to it.
# The default installation path is usually the one below. If yours is different, update this path.
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# NEW: Dictionary to keep track of active users
ACTIVE_USERS = {}

# UPDATED: Setup TWO persistent CSV log files
REGISTRATION_LOG = 'registration_log.csv'
if not os.path.exists(REGISTRATION_LOG):
    with open(REGISTRATION_LOG, 'w', encoding='utf-8') as f:
        f.write("Timestamp,Email,IP_Address\n")

USAGE_LOG = 'usage_log.csv'
if not os.path.exists(USAGE_LOG):
    with open(USAGE_LOG, 'w', encoding='utf-8') as f:
        f.write("Timestamp,Email,Session_ID,IP_Address\n")

# NEW: Setup simple text file for storing access requests
REQUESTS_FILE = 'access_requests.txt'
if not os.path.exists(REQUESTS_FILE):
    with open(REQUESTS_FILE, 'w', encoding='utf-8') as f:
        pass

# NEW: Setup simple text file for storing approved users dynamically
APPROVED_USERS_FILE = 'approved_users.txt'
if not os.path.exists(APPROVED_USERS_FILE):
    with open(APPROVED_USERS_FILE, 'w', encoding='utf-8') as f:
        f.write('haque.mazharul@ap.averydennison.com\n') # Admin is approved by default

# -------------------------------------------------------------------
# FRONTEND: HTML / CSS (Tailwind) / JS (PDF.js)
# -------------------------------------------------------------------
HTML_TEMPLATE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Web PDF Editor</title>
    <script src="https://cdn.tailwindcss.com"></script>
    <script src="https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.min.js"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0/css/all.min.css" rel="stylesheet">
</head>
<body class="bg-gray-100 flex h-screen overflow-hidden font-sans">

    <!-- Left Sidebar (Tools & Upload) -->
    <aside class="w-80 shrink-0 bg-white border-r flex flex-col p-5 shadow-lg z-10 relative overflow-y-auto">
        <h1 class="text-2xl font-bold text-indigo-600 mb-6 flex items-center gap-2">
            <i class="fa-solid fa-file-pdf"></i> DocHub <span class="bg-indigo-100 text-indigo-800 text-xs font-semibold px-2 py-1 rounded ml-auto">v2.2</span>
        </h1>
        
        <!-- Upload Zone (Edit Document) -->
        <div class="mb-6 bg-indigo-600 p-5 rounded-xl border border-indigo-700 shadow-md">
            <h2 class="text-sm font-bold text-white uppercase tracking-wider mb-3 border-b border-indigo-400 pb-2"><i class="fa-solid fa-file-signature mr-2"></i>Edit Document</h2>
            <!-- FIXED: Changed background to non-white bg-indigo-700 and text-white for premium look and readability -->
            <input type="file" id="upload-input" accept="application/pdf" 
                class="block w-full text-xs text-indigo-100 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-indigo-700 file:text-white hover:file:bg-indigo-800 cursor-pointer transition"/>
        </div>

        <!-- Action Tools (Hidden until a file is uploaded) -->
        <div id="tools-section" class="hidden flex-col gap-3 mb-6 border-b pb-6">
            <h2 class="text-xs font-bold text-gray-400 uppercase tracking-wider mb-2">Edit Current Page</h2>
            
            <button onclick="rotatePage()" class="flex items-center gap-3 bg-gray-50 hover:bg-gray-100 text-gray-700 py-2.5 px-4 rounded border transition">
                <i class="fa-solid fa-rotate-right text-indigo-500 w-4"></i> Rotate 90°
            </button>
            
            <!-- MOVED & RECOLORED: Insert PDF Tool -->
            <div class="p-4 border rounded bg-red-50/50 border-red-200 relative">
                <h3 class="text-xs font-bold text-red-700 uppercase tracking-wider mb-3">Insert PDF Here</h3>
                <p class="text-[10px] text-gray-600 mb-2">Inserts a new PDF file immediately <b>after</b> the page you are currently viewing.</p>
                <input type="file" id="insert-pdf-input" accept="application/pdf" class="w-full text-xs text-gray-600 file:mr-2 file:py-1.5 file:px-3 file:rounded file:border-0 file:text-xs file:font-semibold file:bg-red-100 file:text-red-700 hover:file:bg-red-200 mb-2 cursor-pointer transition"/>
                <button onclick="insertPdf()" class="w-full flex items-center justify-center gap-2 bg-white hover:bg-red-50 text-red-700 font-medium py-2 px-4 rounded border border-red-200 shadow-sm transition">
                    <i class="fa-solid fa-file-import"></i> Insert File
                </button>
            </div>

            <button onclick="deletePage()" class="flex items-center gap-3 bg-red-50 hover:bg-red-100 text-red-700 py-2.5 px-4 rounded border border-red-200 transition">
                <i class="fa-solid fa-trash-can w-4"></i> Delete Page
            </button>

            <!-- NEW: Extract Text Button (Added Padlock Icon) -->
            <button onclick="extractText()" class="flex items-center gap-3 bg-blue-50 hover:bg-blue-100 text-blue-700 py-2.5 px-4 rounded border border-blue-200 transition relative">
                <i class="fa-solid fa-copy w-4"></i> Copy Page Text
                <i class="fa-solid fa-lock text-[10px] text-blue-400 absolute right-4" title="Requires Approval"></i>
            </button>

            <!-- Text Replacement Tool (Added Padlock Icon) -->
            <div class="mt-4 p-4 border rounded bg-indigo-50/50 relative">
                <i class="fa-solid fa-lock text-[10px] text-indigo-400 absolute top-4 right-4" title="Requires Approval"></i>
                <h3 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-3">Replace Text</h3>
                <input type="text" id="find-text" placeholder="Word to find..." class="w-full text-sm p-2 border rounded border-gray-300 mb-2" />
                <input type="text" id="replace-text" placeholder="Replace with..." class="w-full text-sm p-2 border rounded border-gray-300 mb-2" />
                <input type="number" id="occurrence-index" placeholder="Instance (leave blank for ALL)" min="1" title="If the word appears 3 times and you want to replace the 2nd one, type 2. Leave blank to replace all." class="w-full text-sm p-2 border rounded border-gray-300 mb-2" />
                
                <!-- New Tip for Spacing -->
                <div class="bg-blue-50 border border-blue-100 rounded p-2 mb-3">
                    <p class="text-[10px] text-blue-700 leading-tight">
                        <i class="fa-solid fa-circle-info mr-1"></i><b>Spacing Tip:</b> PDFs don't auto-flow. To prevent awkward gaps when replacing words of different lengths, replace the surrounding words too.<br>
                        <i>Ex: Find "Mymensingh, PO" -> Replace with "Gazipur, PO"</i>
                    </p>
                </div>

                <button onclick="replaceText()" class="w-full flex items-center justify-center gap-2 bg-white hover:bg-indigo-50 text-indigo-700 font-medium py-2 px-4 rounded border border-indigo-200 shadow-sm transition">
                    <i class="fa-solid fa-pen-to-square"></i> Apply to Page
                </button>
            </div>

            <!-- NEW: Add Signature Tool -->
            <div class="mt-4 p-4 border rounded bg-emerald-50/50 relative">
                <h3 class="text-xs font-bold text-gray-500 uppercase tracking-wider mb-3">Add Signature</h3>
                <input type="file" id="signature-upload" accept=".png,.jpg,.jpeg" class="w-full text-xs text-gray-500 file:mr-2 file:py-1.5 file:px-3 file:rounded file:border-0 file:text-xs file:font-semibold file:bg-emerald-100 file:text-emerald-700 hover:file:bg-emerald-200 mb-2 cursor-pointer transition"/>
                <p class="text-[10px] text-gray-500 mb-1" id="sig-instruction">Upload a signature image (.png or .jpg), then click anywhere on the document to stamp it.</p>
            </div>
        </div>

        <!-- Merge PDFs Zone (Thin Green Layout) -->
        <div class="mb-6 bg-emerald-50/50 p-5 rounded-xl border border-emerald-200 shadow-sm">
            <h2 class="text-sm font-bold text-emerald-800 uppercase tracking-wider mb-3 border-b border-emerald-200 pb-2"><i class="fa-solid fa-layer-group mr-2"></i>Merge PDFs</h2>
            <!-- FIXED: Changed background to non-white bg-emerald-200 and text-emerald-800 for elegant theme layout -->
            <input type="file" id="merge-input" accept="application/pdf" multiple 
                class="block w-full text-xs text-emerald-800 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-emerald-200 file:text-emerald-800 hover:file:bg-emerald-300 cursor-pointer transition mb-4"/>
            <button onclick="mergePdfs()" class="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-bold py-2.5 px-3 rounded-lg shadow-sm transition">
                <i class="fa-solid fa-layer-group"></i> Merge Files
            </button>
        </div>

        <!-- Compress PDF Zone (Thin Green Layout) -->
        <div class="mb-6 bg-emerald-50/50 p-5 rounded-xl border border-emerald-200 shadow-sm">
            <h2 class="text-sm font-bold text-emerald-800 uppercase tracking-wider mb-3 border-b border-emerald-200 pb-2"><i class="fa-solid fa-compress mr-2"></i>Compress PDF</h2>
            <!-- FIXED: Standardized file select styles to contrast thin green theme -->
            <input type="file" id="compress-input" accept="application/pdf" 
                class="block w-full text-xs text-emerald-800 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-emerald-200 file:text-emerald-800 hover:file:bg-emerald-300 cursor-pointer transition mb-3"/>
            
            <select id="compress-mode" class="w-full text-sm p-2.5 border rounded-lg border-emerald-200 mb-4 text-gray-800 bg-white shadow-sm focus:ring-2 focus:ring-emerald-500 focus:outline-none">
                <option value="auto_1mb" selected>Auto Compress (Target: ~1 MB)</option>
                <option value="auto_512kb">Auto Compress (Target: ~512 KB)</option>
            </select>

            <button onclick="compressPdf()" class="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-bold py-2.5 px-3 rounded-lg shadow-sm transition">
                <i class="fa-solid fa-compress"></i> Compress PDF
            </button>
        </div>

        <!-- Unzip File Zone (Thin Green Layout) -->
        <div class="mb-6 bg-emerald-50/50 p-5 rounded-xl border border-emerald-200 shadow-sm">
            <h2 class="text-sm font-bold text-emerald-800 uppercase tracking-wider mb-3 border-b border-emerald-200 pb-2"><i class="fa-solid fa-file-zipper mr-2"></i>Unzip File</h2>
            <!-- FIXED: Standardized file select styles to contrast thin green theme -->
            <input type="file" id="unzip-input" accept=".zip" 
                class="block w-full text-xs text-emerald-800 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-emerald-200 file:text-emerald-800 hover:file:bg-emerald-300 cursor-pointer transition mb-4"/>
            <button onclick="unzipFile()" class="w-full flex items-center justify-center gap-2 bg-emerald-600 hover:bg-emerald-700 text-white text-sm font-bold py-2.5 px-3 rounded-lg shadow-sm transition">
                <i class="fa-solid fa-file-zipper"></i> Unzip Contents
            </button>
        </div>

        <!-- ================= UNIVERSAL CONVERTER ================= -->
        <div class="mb-6 bg-blue-600 p-5 rounded-xl border border-blue-700 shadow-md">
            <h2 class="text-sm font-bold text-white uppercase tracking-wider mb-4 border-b border-blue-400 pb-2"><i class="fa-solid fa-right-left mr-2"></i>Document Converter</h2>
            
            <label class="block text-xs font-semibold text-blue-100 mb-1">Select Conversion Type</label>
            <select id="convert-type" class="w-full text-sm p-2.5 border rounded-lg border-blue-500 mb-4 text-gray-800 bg-white shadow-sm focus:ring-2 focus:ring-white focus:outline-none">
                <optgroup label="Convert FROM PDF">
                    <option value="pdf_to_word">PDF to Word</option>
                    <option value="pdf_to_excel">PDF to Excel</option>
                    <option value="pdf_to_jpg">PDF to JPG</option>
                </optgroup>
                <optgroup label="Convert TO PDF & Word">
                    <option value="word_to_pdf">Word to PDF</option>
                    <option value="excel_to_pdf">Excel to PDF</option>
                    <option value="jpg_to_pdf">JPG to PDF (Supports Multiple)</option>
                    <option value="jpg_to_word">JPG to Word (OCR)</option>
                </optgroup>
            </select>

            <label class="block text-xs font-semibold text-blue-100 mb-1">Upload File(s)</label>
            <!-- FIXED: Changed background to non-white bg-blue-700 for high readability on conversion panel -->
            <input type="file" id="universal-convert-input" accept=".pdf,.doc,.docx,.xls,.xlsx,.jpg,.jpeg,.png" multiple
                class="block w-full text-xs text-blue-100 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-blue-700 file:text-white hover:file:bg-blue-800 cursor-pointer transition mb-4"/>
            
            <button onclick="executeConversion()" class="w-full flex items-center justify-center gap-2 bg-white hover:bg-blue-50 text-blue-700 text-sm font-bold py-2.5 px-3 rounded-lg shadow-sm transition">
                <i class="fa-solid fa-wand-magic-sparkles"></i> Convert File
            </button>
        </div>
        
        <!-- NEW: Create Secure PDF Zone -->
        <div class="mb-6 bg-blue-600 p-5 rounded-xl border border-blue-700 shadow-md">
            <h2 class="text-sm font-bold text-white uppercase tracking-wider mb-2 border-b border-blue-400 pb-2"><i class="fa-solid fa-file-shield mr-2"></i>Create Read-Only PDF <i class="fa-solid fa-lock ml-1 text-blue-200 text-xs"></i></h2>
            <p class="text-[10px] text-blue-100 mb-3">Convert Word, JPG, or PDF into a secured, non-editable file.</p>
            <!-- FIXED: Changed background to non-white bg-blue-700 for high readability on security panel -->
            <input type="file" id="secure-pdf-input" accept=".pdf,.jpg,.jpeg,.png,.doc,.docx"
                class="block w-full text-xs text-blue-100 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:font-semibold file:bg-blue-700 file:text-white hover:file:bg-blue-400 cursor-pointer transition mb-4"/>
            <button onclick="createSecurePdf()" class="w-full flex items-center justify-center gap-2 bg-white hover:bg-blue-50 text-blue-700 text-sm font-bold py-2.5 px-3 rounded-lg shadow-sm transition">
                <i class="fa-solid fa-shield-halved"></i> Lock & Download
            </button>
        </div>
        
        <!-- Save & Download Action -->
        <div id="download-section" class="hidden mt-auto pt-2 pb-4">
            <button onclick="downloadPdf()" class="w-full flex items-center justify-center gap-2 bg-red-600 hover:bg-red-700 text-white font-bold py-3 px-4 rounded shadow-md transition">
                <i class="fa-solid fa-download"></i> Save & Download
            </button>
        </div>
        
        <!-- UPDATED: Download Access Logs Buttons (Now triggers password prompt) -->
        <div class="mt-auto pt-4 text-center border-t border-gray-100 flex flex-col gap-2">
            <button onclick="downloadLogs('registration')" class="text-xs text-gray-400 hover:text-indigo-600 transition" title="Download first-time registration records">
                <i class="fa-solid fa-lock text-xs mr-1"></i> Admin: Registration Logs
            </button>
            <button onclick="downloadLogs('usage')" class="text-xs text-gray-400 hover:text-indigo-600 transition" title="Download daily portal usage records">
                <i class="fa-solid fa-lock text-xs mr-1"></i> Admin: Usage Logs
            </button>
        </div>

        <!-- Loading Overlay for sidebar -->
        <div id="loading" class="hidden absolute inset-0 bg-white/80 flex items-center justify-center z-20">
            <div class="animate-spin rounded-full h-8 w-8 border-b-2 border-indigo-600"></div>
        </div>
    </aside>

    <!-- Main Workspace (Viewer) -->
    <main class="flex-1 flex flex-col bg-gray-200">
        <!-- Top Navigation / Toolbar -->
        <header class="h-14 bg-white shadow-sm flex items-center justify-between px-6 z-0">
            <div id="doc-info" class="text-sm font-medium text-gray-600">No document loaded</div>
            
            <!-- NEW: Live Users Badge (Hidden by default via inline style) -->
            <div class="flex items-center gap-3 ml-auto mr-4" id="live-users-container" style="display: none;">
                
                <!-- NEW: Admin Notification Bell (For Access Requests) -->
                <button id="admin-bell" onclick="showRequests()" class="hidden relative items-center justify-center w-8 h-8 rounded-full bg-red-50 text-red-600 hover:bg-red-100 transition mr-2">
                    <i class="fa-solid fa-bell"></i>
                    <span id="admin-bell-count" class="absolute -top-1 -right-1 bg-red-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-full shadow">0</span>
                </button>

                <!-- UPDATED: Made the badge clickable to show all users -->
                <div onclick="showOnlineUsers()" class="hidden md:flex items-center gap-2 text-xs font-medium text-green-600 bg-green-50 px-3 py-1.5 rounded-full border border-green-200 shadow-sm cursor-pointer hover:bg-green-100 transition" title="Click to see full list of online users">
                    <span class="relative flex h-2 w-2">
                      <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-green-400 opacity-75"></span>
                      <span class="relative inline-flex rounded-full h-2 w-2 bg-green-500"></span>
                    </span>
                    <span id="online-users-count">1 Online</span>
                    <span id="online-users-list" class="text-gray-500 font-normal ml-1 truncate max-w-[200px]">(Connecting...)</span>
                </div>
            </div>

            <!-- Pagination Controls -->
            <div id="pagination" class="hidden flex items-center gap-4">
                <button onclick="changePage(-1)" class="p-2 rounded-full hover:bg-gray-100 disabled:opacity-30 transition" id="btn-prev">
                    <i class="fa-solid fa-chevron-left"></i>
                </button>
                <span class="text-sm font-medium text-gray-700">Page <span id="page-num">0</span> of <span id="page-count">0</span></span>
                <button onclick="changePage(1)" class="p-2 rounded-full hover:bg-gray-100 disabled:opacity-30 transition" id="btn-next">
                    <i class="fa-solid fa-chevron-right"></i>
                </button>
            </div>
        </header>

        <!-- Canvas Container -->
        <div class="flex-1 overflow-auto p-8 flex justify-center items-start">
            <div id="canvas-container" class="relative inline-block shadow-2xl bg-white hidden transition-all">
                <canvas id="pdf-render" class="block"></canvas>
                
                <!-- NEW: Floating Signature Overlay (Draggable & Resizable) -->
                <div id="sig-overlay" class="hidden absolute border-2 border-dashed border-blue-500 cursor-move z-50 shadow-md bg-white/10 backdrop-blur-sm" style="top: 50px; left: 50px;">
                    <img id="sig-image" src="" class="pointer-events-none select-none" style="width: 150px; height: auto; min-width: 50px;" />
                    
                    <!-- Resize Handle -->
                    <div id="sig-resize-handle" class="absolute -right-2 -bottom-2 w-4 h-4 bg-blue-600 rounded-full cursor-se-resize border-2 border-white shadow hover:scale-125 transition-transform"></div>
                    
                    <!-- Action Buttons -->
                    <div class="absolute -bottom-12 left-1/2 transform -translate-x-1/2 bg-white shadow-xl rounded-full flex gap-1 p-1.5 border border-gray-200 cursor-default pointer-events-auto w-max">
                        <button onclick="applySignature()" class="px-3 py-1 bg-green-500 hover:bg-green-600 text-white rounded-full text-xs font-bold transition flex items-center gap-1 shadow-sm"><i class="fa-solid fa-check"></i> Apply</button>
                        <button onclick="cancelSignature()" class="px-3 py-1 bg-red-100 hover:bg-red-200 text-red-700 rounded-full text-xs font-bold transition flex items-center gap-1 shadow-sm"><i class="fa-solid fa-xmark"></i></button>
                    </div>
                </div>
            </div>
            
            <!-- Empty State -->
            <div id="empty-state" class="h-full flex flex-col items-center justify-center text-gray-400">
                <i class="fa-regular fa-file-pdf text-6xl mb-4 text-gray-300"></i>
                <p class="text-lg font-medium">Upload a PDF to start editing</p>
            </div>
            
            <!-- Extracted Files State -->
            <div id="extracted-container" class="hidden w-full max-w-3xl bg-white rounded-lg shadow-lg p-8"></div>
        </div>
    </main>

    <!-- NEW: Text Extraction Modal Popup -->
    <div id="text-modal" class="hidden fixed inset-0 bg-black/50 flex items-center justify-center z-50">
        <div class="bg-white rounded-lg shadow-xl w-full max-w-2xl flex flex-col overflow-hidden">
            <div class="px-6 py-4 border-b flex justify-between items-center bg-gray-50">
                <h3 class="font-bold text-gray-800">Extracted Page Text</h3>
                <button onclick="closeTextModal()" class="text-gray-500 hover:text-gray-800"><i class="fa-solid fa-xmark"></i></button>
            </div>
            <div class="p-6 flex-1">
                <p class="text-xs text-gray-500 mb-2"><i class="fa-solid fa-circle-info mr-1"></i> You can easily copy sentences from the box below. If the box is empty, your document is a scanned image and you should use the "PDF to Word" OCR tool instead.</p>
                <textarea id="extracted-text-area" class="w-full h-64 p-3 border rounded focus:ring-2 focus:ring-indigo-500 focus:outline-none font-mono text-sm resize-none"></textarea>
            </div>
            <div class="px-6 py-4 border-t bg-gray-50 flex justify-end gap-3">
                <button onclick="closeTextModal()" class="px-4 py-2 text-gray-600 font-medium hover:bg-gray-200 rounded transition">Close</button>
                <button onclick="copyExtractedText()" class="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-medium rounded transition flex items-center gap-2">
                    <i class="fa-solid fa-copy"></i> Copy
                </button>
            </div>
        </div>
    </div>

    <!-- UPDATED: User Name Prompt Modal for Live Tracking -->
    <div id="name-modal" class="hidden fixed inset-0 bg-gray-900/80 backdrop-blur-sm flex items-center justify-center z-[100]">
        <div class="bg-white rounded-xl shadow-2xl p-8 max-w-sm w-full transform transition-all text-center">
            <div class="w-16 h-16 bg-indigo-100 text-indigo-600 rounded-full flex items-center justify-center text-3xl mx-auto mb-4 shadow-inner">
                <i class="fa-solid fa-user-shield"></i>
            </div>
            <h2 class="text-xl font-bold text-gray-800 mb-2">Corporate Login</h2>
            <p class="text-sm text-gray-500 mb-6">Please verify your Avery Dennison employee email to access this portal.</p>
            <input type="email" id="user-name-input" placeholder="name@ap.averydennison.com" class="w-full px-4 py-3 border-2 border-gray-200 rounded-lg focus:ring-4 focus:ring-indigo-100 focus:border-indigo-500 transition-all text-center font-medium text-gray-700 mb-4 outline-none" autocomplete="email" onkeypress="if(event.key === 'Enter') registerUser()">
            <button onclick="registerUser()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-4 rounded-lg shadow-md transition-colors">
                Verify & Enter
            </button>
        </div>
    </div>

    <script>
        // Setup PDF.js web worker
        pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.16.105/pdf.worker.min.js';

        // NEW: Universal Save Helper (Prompts user for folder location using File System Access API)
        async function saveFileToDisk(fetchUrl, suggestedName, mimeType, ext, desc) {
            if (window.showSaveFilePicker) {
                try {
                    const handle = await window.showSaveFilePicker({
                        suggestedName: suggestedName,
                        types: [{
                            description: desc,
                            accept: {[mimeType]: [ext]},
                        }],
                    });
                    
                    loading.classList.remove('hidden');
                    const writable = await handle.createWritable();
                    const response = await fetch(fetchUrl);
                    const blob = await response.blob();
                    await writable.write(blob);
                    await writable.close();
                    
                    alert('File saved successfully to your chosen folder!');
                } catch (err) {
                    if (err.name !== 'AbortError') {
                        console.error(err);
                        alert('Save failed, falling back to standard download.');
                        window.location.href = fetchUrl; // Fallback to classic browser download
                    }
                } finally {
                    loading.classList.add('hidden');
                }
            } else {
                // Fallback for older browsers
                window.location.href = fetchUrl;
            }
        }

        // UPDATED: Live User Tracking Logic
        let mySessionId = 'sess_' + Math.random().toString(36).substr(2, 9) + Date.now();
        let myName = localStorage.getItem('dochub_email'); 
        let APPROVED_USERS = [];

        window.addEventListener('DOMContentLoaded', () => {
            if (!myName) {
                document.getElementById('name-modal').classList.remove('hidden');
            } else {
                startHeartbeat();
            }
        });

        function registerUser() {
            const emailInput = document.getElementById('user-name-input').value.trim().toLowerCase();
            if (!emailInput || !emailInput.endsWith('@ap.averydennison.com')) {
                alert("Access Denied: Please enter a valid @ap.averydennison.com email address.");
                return;
            }
            myName = emailInput;
            localStorage.setItem('dochub_email', myName);
            document.getElementById('name-modal').classList.add('hidden');
            fetch('/register_user', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ email: myName })
            });
            startHeartbeat();
        }

        function startHeartbeat() {
            sendHeartbeat();
            setInterval(sendHeartbeat, 10000); 
        }

        let currentRequests = [];
        let currentOnlineUsers = [];

        async function sendHeartbeat() {
            try {
                const response = await fetch('/heartbeat', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ session_id: mySessionId, name: myName })
                });
                const data = await response.json();
                if (data.approved_users) {
                    APPROVED_USERS = data.approved_users;
                }
                if (data.users) {
                    updateOnlineUsersUI(data.users, data.pending_requests);
                }
            } catch (err) {
                console.log("Heartbeat disconnected.");
            }
        }

        function updateOnlineUsersUI(users, pending_requests) {
            if (myName === 'haque.mazharul@ap.averydennison.com') {
                document.getElementById('live-users-container').style.display = 'flex';
                if (pending_requests && pending_requests.length > 0) {
                    document.getElementById('admin-bell').classList.remove('hidden');
                    document.getElementById('admin-bell').classList.add('flex');
                    document.getElementById('admin-bell-count').textContent = pending_requests.length;
                    currentRequests = pending_requests;
                } else {
                    document.getElementById('admin-bell').classList.add('hidden');
                    document.getElementById('admin-bell').classList.remove('flex');
                    currentRequests = [];
                }
            } else {
                document.getElementById('live-users-container').style.display = 'none';
                return; 
            }
            const countSpan = document.getElementById('online-users-count');
            const listSpan = document.getElementById('online-users-list');
            countSpan.textContent = `${users.length} Online`;
            let displayNames = users.map(u => {
                if (u === myName) return 'You';
                return u.split('@')[0].replace('.', ' '); 
            });
            currentOnlineUsers = displayNames;
            listSpan.textContent = `(${displayNames.join(', ')})`;
        }

        function showOnlineUsers() {
            if (currentOnlineUsers.length === 0) return;
            alert("🟢 Currently Online (" + currentOnlineUsers.length + "):\n\n• " + currentOnlineUsers.join("\n• "));
        }

        function downloadLogs(logType) {
            const password = prompt(`Admin Area: Please enter the password to download the ${logType} logs.`);
            if (password) {
                window.location.href = `/download_logs?type=${logType}&key=${encodeURIComponent(password)}`;
            }
        }

        function showRequests() {
            if (currentRequests.length === 0) return;
            const msg = "The following users are requesting feature access:\n\n• " + currentRequests.join('\n• ') + "\n\nClick 'OK' to instantly approve these users!";
            if(confirm(msg)) {
                fetch('/approve_users', { method: 'POST' });
                document.getElementById('admin-bell-count').textContent = '0';
                document.getElementById('admin-bell').classList.add('hidden');
                currentRequests = [];
                alert("Users have been approved successfully!");
            }
        }

        let currentPdf = null;
        let pageNum = 1;
        let fileId = null;
        let totalPages = 0;
        let signatureFile = null;
        let isSignatureMode = false;

        const canvas = document.getElementById('pdf-render');
        const ctx = canvas.getContext('2d');
        const toolsSection = document.getElementById('tools-section');
        const downloadSection = document.getElementById('download-section');
        const emptyState = document.getElementById('empty-state');
        const canvasContainer = document.getElementById('canvas-container');
        const extractedContainer = document.getElementById('extracted-container');
        const pagination = document.getElementById('pagination');
        const loading = document.getElementById('loading');
        const docInfo = document.getElementById('doc-info');
        const sigOverlay = document.getElementById('sig-overlay');
        const sigImage = document.getElementById('sig-image');
        const sigResizeHandle = document.getElementById('sig-resize-handle');

        document.getElementById('upload-input').addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file || file.type !== 'application/pdf') {
                alert('Please select a valid PDF file.');
                return;
            }
            docInfo.textContent = file.name;
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file', file);
            try {
                const response = await fetch('/upload', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.file_id) {
                    fileId = data.file_id;
                    pageNum = 1;
                    await loadPdf();
                    toolsSection.classList.remove('hidden');
                    toolsSection.classList.add('flex');
                    downloadSection.classList.remove('hidden');
                    emptyState.classList.add('hidden');
                    canvasContainer.classList.remove('hidden');
                    pagination.classList.remove('hidden');
                }
            } catch (err) {
                alert("Upload failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        });

        async function mergePdfs() {
            const files = document.getElementById('merge-input').files;
            if (files.length < 2) {
                alert('Please select at least 2 PDF files to merge (hold Ctrl/Cmd to select multiple files).');
                return;
            }
            loading.classList.remove('hidden');
            const formData = new FormData();
            for (let i = 0; i < files.length; i++) {
                formData.append('files[]', files[i]);
            }
            try {
                const response = await fetch('/merge', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.file_id) {
                    fileId = data.file_id;
                    pageNum = 1;
                    docInfo.textContent = "Merged_Document.pdf";
                    await loadPdf();
                    toolsSection.classList.remove('hidden');
                    toolsSection.classList.add('flex');
                    downloadSection.classList.remove('hidden');
                    emptyState.classList.add('hidden');
                    canvasContainer.classList.remove('hidden');
                    pagination.classList.remove('hidden');
                    document.getElementById('merge-input').value = "";
                    document.getElementById('upload-input').value = "";
                    setTimeout(() => {
                        alert(`Success! Files merged into a ${totalPages}-page document.\n\nUse the < and > arrows at the top right to see the rest of your document.`);
                    }, 300);
                } else {
                    alert(data.error || "Merge failed.");
                }
            } catch (err) {
                alert("Merge failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        // UPDATED: Now uses Save File Picker Dialog
        async function compressPdf() {
            const fileInput = document.getElementById('compress-input');
            const file = fileInput.files[0];
            const mode = document.getElementById('compress-mode').value;
            if (!file) {
                alert('Please select a PDF file to compress.');
                return;
            }
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file', file);
            formData.append('mode', mode); 
            try {
                const response = await fetch('/compress_pdf', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.file_id) {
                    const finalSizeKB = Math.round(data.size_kb);
                    alert(`Compression complete! Final size is ${finalSizeKB} KB. Choose where to save it next.`);
                    await saveFileToDisk(`/download_compressed/${data.file_id}`, `compressed_${file.name}`, 'application/pdf', '.pdf', 'PDF Document');
                    fileInput.value = "";
                } else {
                    alert(data.error || "Compression failed.");
                }
            } catch (err) {
                alert("Compression failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        // NEW: Consolidated Universal Conversion Handler
        async function executeConversion() {
            const type = document.getElementById('convert-type').value;
            const fileInput = document.getElementById('universal-convert-input');
            const files = fileInput.files;

            if (files.length === 0) {
                alert('Please select at least one file to convert.');
                return;
            }

            const config = {
                'pdf_to_word': { url: '/convert_pdf_to_word', multi: false, dl_url: '/download_word/', ext: '.docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', desc: 'Word Document', replaceExt: '.docx' },
                'pdf_to_excel': { url: '/convert_pdf_to_excel', multi: false, dl_url: '/download_excel/', ext: '.xlsx', mime: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', desc: 'Excel Document', replaceExt: '.xlsx' },
                'pdf_to_jpg': { url: '/convert_pdf_to_jpg', multi: false, dl_url: '/download_jpg/', ext: '.jpg', mime: 'image/jpeg', desc: 'JPEG Image', replaceExt: '.jpg' },
                'word_to_pdf': { url: '/convert_word_to_pdf', multi: false, render: true, newExt: '.pdf' },
                'excel_to_pdf': { url: '/convert_excel', multi: false, render: true, newExt: '.pdf' },
                'jpg_to_pdf': { url: '/convert_jpg_to_pdf', multi: true, render: true, newExt: '.pdf' },
                'jpg_to_word': { url: '/convert_jpg_to_word', multi: false, dl_url: '/download_word/', ext: '.docx', mime: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', desc: 'Word Document', replaceExt: '.docx' }
            }[type];

            const file = files[0];
            const filename = file.name.toLowerCase();
            
            // Basic validation check
            if (type.startsWith('pdf_') && !filename.endsWith('.pdf')) {
                 return alert('Please select a PDF file for this conversion type.');
            }
            if (type === 'word_to_pdf' && !filename.match(/\.docx?$/)) {
                 return alert('Please select a Word file (.doc, .docx).');
            }
            if (type === 'excel_to_pdf' && !filename.match(/\.xlsx?$/)) {
                 return alert('Please select an Excel file (.xls, .xlsx).');
            }
            if (type.startsWith('jpg_')) {
                 for(let i=0; i<files.length; i++) {
                     if(!files[i].name.toLowerCase().match(/\.(jpg|jpeg|png)$/)) {
                         return alert('Please select image files (.jpg, .jpeg, .png).');
                     }
                 }
            }

            loading.classList.remove('hidden');
            const formData = new FormData();

            if (config.multi) {
                for (let i = 0; i < files.length; i++) {
                    formData.append('files[]', files[i]);
                }
            } else {
                formData.append('file', file);
            }

            try {
                const response = await fetch(config.url, { method: 'POST', body: formData });
                const data = await response.json();
                
                if (data.file_id) {
                    if (config.render) {
                        // Render in document viewer (Word to PDF, Excel to PDF, JPG to PDF)
                        fileId = data.file_id;
                        pageNum = 1;
                        if (config.multi && files.length > 1) {
                            docInfo.textContent = `Merged_${files.length}_Images.pdf`;
                        } else {
                            docInfo.textContent = file.name.replace(/\.[^/.]+$/, "") + config.newExt;
                        }
                        await loadPdf();
                        toolsSection.classList.remove('hidden');
                        toolsSection.classList.add('flex');
                        downloadSection.classList.remove('hidden');
                        emptyState.classList.add('hidden');
                        canvasContainer.classList.remove('hidden');
                        pagination.classList.remove('hidden');
                    } else {
                        // Direct Download behavior (PDF to Word/Excel/JPG, JPG to Word)
                        const suggestedName = file.name.replace(/\.[^/.]+$/, "") + config.replaceExt;
                        await saveFileToDisk(`${config.dl_url}${data.file_id}`, suggestedName, config.mime, config.ext, config.desc);
                    }
                    fileInput.value = "";
                } else {
                    let errMsg = data.error || "Conversion failed.";
                    if (type === 'pdf_to_excel' && !data.error) errMsg += " Ensure the PDF contains structured tables.";
                    alert(errMsg);
                }
            } catch (err) {
                alert("Conversion failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        // NEW: Secure PDF Javascript Handler
        async function createSecurePdf() {
            const fileInput = document.getElementById('secure-pdf-input');
            const file = fileInput.files[0];
            if (!file) {
                alert('Please select a file to secure (.jpg, .doc, .pdf).');
                return;
            }
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file', file);
            try {
                const response = await fetch('/create_secure_pdf', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.file_id) {
                    const suggestedName = "Locked_" + file.name.replace(/\.[^/.]+$/, "") + ".pdf";
                    await saveFileToDisk(`/download/${data.file_id}`, suggestedName, 'application/pdf', '.pdf', 'Secure PDF Document');
                    fileInput.value = "";
                } else {
                    alert(data.error || "Securing failed.");
                }
            } catch (err) {
                alert("Action failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        async function unzipFile() {
            const fileInput = document.getElementById('unzip-input');
            const file = fileInput.files[0];
            if (!file) {
                alert('Please select a ZIP file to unzip.');
                return;
            }
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file', file);
            try {
                const response = await fetch('/unzip_file', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.extract_id) {
                    canvasContainer.classList.add('hidden');
                    emptyState.classList.add('hidden');
                    pagination.classList.add('hidden');
                    toolsSection.classList.add('hidden');
                    downloadSection.classList.add('hidden');
                    docInfo.textContent = "Unzipped Files: " + file.name;
                    extractedContainer.classList.remove('hidden');
                    let html = `
                        <div class="flex justify-between items-center mb-4 border-b pb-3">
                            <h2 class="text-xl font-bold text-gray-800">Contents of ${file.name}</h2>
                            <button onclick="downloadNativeFolder('${data.extract_id}', '${encodeURIComponent(JSON.stringify(data.files))}')" 
                               class="bg-blue-600 hover:bg-blue-700 text-white px-4 py-2 rounded shadow text-sm font-bold transition flex items-center gap-2">
                                <i class="fa-solid fa-folder-arrow-down"></i> Save to Local Folder
                            </button>
                        </div>
                        <ul class="divide-y divide-gray-200 max-h-[60vh] overflow-y-auto">
                    `;
                    data.files.forEach(f => {
                        html += `
                            <li class="py-3 flex justify-between items-center">
                                <span class="text-sm font-medium text-gray-700 break-all mr-4"><i class="fa-regular fa-file mr-2 text-gray-400"></i>${f}</span>
                                <a href="/download_extracted/${data.extract_id}/${encodeURIComponent(f)}" download
                                   class="shrink-0 bg-yellow-100 text-yellow-700 hover:bg-yellow-200 px-3 py-1 rounded text-sm font-semibold transition flex items-center gap-2">
                                    <i class="fa-solid fa-download"></i> Download
                                </a>
                            </li>
                        `;
                    });
                    if(data.files.length === 0) {
                        html += '<li class="py-3 text-sm text-gray-500">The ZIP archive was empty.</li>';
                    }
                    html += '</ul>';
                    extractedContainer.innerHTML = html;
                    fileInput.value = "";
                } else {
                    alert(data.error || "Unzipping failed.");
                }
            } catch (err) {
                alert("Unzipping failed. Make sure the backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        async function loadPdf() {
            const url = `/document/${fileId}?t=${new Date().getTime()}`; 
            const loadingTask = pdfjsLib.getDocument(url);
            currentPdf = await loadingTask.promise;
            totalPages = currentPdf.numPages;
            document.getElementById('page-count').textContent = totalPages;
            if (extractedContainer) extractedContainer.classList.add('hidden');
            renderPage(pageNum);
        }

        async function renderPage(num) {
            const page = await currentPdf.getPage(num);
            const viewport = page.getViewport({ scale: 1.5 });
            canvas.height = viewport.height;
            canvas.width = viewport.width;
            const renderContext = { canvasContext: ctx, viewport: viewport };
            await page.render(renderContext).promise;
            document.getElementById('page-num').textContent = num;
            document.getElementById('btn-prev').disabled = num <= 1;
            document.getElementById('btn-next').disabled = num >= totalPages;
        }

        function changePage(offset) {
            if (pageNum + offset < 1 || pageNum + offset > totalPages) return;
            pageNum += offset;
            renderPage(pageNum);
        }

        async function deletePage() {
            if (!fileId || totalPages <= 1) {
                alert("Cannot delete the only page of a document.");
                return;
            }
            if(!confirm(`Are you sure you want to permanently delete page ${pageNum}?`)) return;
            loading.classList.remove('hidden');
            try {
                const response = await fetch('/delete_page', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ file_id: fileId, page: pageNum - 1 }) 
                });
                const data = await response.json();
                if (data.success) {
                    if (pageNum > data.total_pages) pageNum = data.total_pages;
                    await loadPdf();
                }
            } finally {
                loading.classList.add('hidden');
            }
        }

        async function rotatePage() {
            if (!fileId) return;
            loading.classList.remove('hidden');
            try {
                const response = await fetch('/rotate_page', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ file_id: fileId, page: pageNum - 1, rotation: 90 }) 
                });
                const data = await response.json();
                if (data.success) await loadPdf();
            } finally {
                loading.classList.add('hidden');
            }
        }

        async function insertPdf() {
            if (!fileId) return;
            const fileInput = document.getElementById('insert-pdf-input');
            const file = fileInput.files[0];
            if (!file) {
                alert("Please select a PDF file to insert.");
                return;
            }
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file_id', fileId);
            formData.append('page', pageNum); 
            formData.append('file', file);
            try {
                const response = await fetch('/insert_pdf', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.success) {
                    await loadPdf(); 
                    fileInput.value = "";
                    setTimeout(() => {
                        alert(`Successfully inserted ${data.inserted_pages} pages into the document!`);
                    }, 300);
                } else {
                    alert("Failed to insert PDF: " + data.error);
                }
            } catch (err) {
                alert("Error connecting to server.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        document.getElementById('signature-upload').addEventListener('change', (e) => {
            signatureFile = e.target.files[0];
            if (signatureFile) {
                const reader = new FileReader();
                reader.onload = function(event) {
                    sigImage.src = event.target.result;
                    sigImage.style.width = '150px';
                    sigOverlay.style.left = '50px';
                    sigOverlay.style.top = '50px';
                    sigOverlay.classList.remove('hidden');
                    document.getElementById('sig-instruction').innerHTML = "<span class='text-emerald-600 font-bold'>Drag and resize!</span> <br>Position your signature, then click Apply.";
                };
                reader.readAsDataURL(signatureFile);
            }
        });

        let isDragging = false;
        let isResizing = false;
        let startX, startY, startLeft, startTop, startWidth;

        sigOverlay.addEventListener('mousedown', (e) => {
            if (e.target === sigResizeHandle) return; 
            if (e.target.closest('button')) return; 
            isDragging = true;
            startX = e.clientX;
            startY = e.clientY;
            startLeft = parseFloat(sigOverlay.style.left) || 0;
            startTop = parseFloat(sigOverlay.style.top) || 0;
            e.preventDefault();
        });

        sigResizeHandle.addEventListener('mousedown', (e) => {
            isResizing = true;
            startX = e.clientX;
            startWidth = parseFloat(sigImage.style.width) || sigImage.clientWidth;
            e.preventDefault();
            e.stopPropagation();
        });

        document.addEventListener('mousemove', (e) => {
            if (isDragging) {
                const dx = e.clientX - startX;
                const dy = e.clientY - startY;
                sigOverlay.style.left = `${startLeft + dx}px`;
                sigOverlay.style.top = `${startTop + dy}px`;
            } else if (isResizing) {
                const dx = e.clientX - startX;
                const newWidth = Math.max(40, startWidth + dx); 
                sigImage.style.width = `${newWidth}px`;
            }
        });

        document.addEventListener('mouseup', () => {
            isDragging = false;
            isResizing = false;
        });

        function cancelSignature() {
            sigOverlay.classList.add('hidden');
            signatureFile = null;
            document.getElementById('signature-upload').value = "";
            document.getElementById('sig-instruction').innerText = "Upload a signature image (.png or .jpg), then click anywhere on the document to stamp it.";
        }

        async function applySignature() {
            if (!signatureFile || !fileId || !currentPdf) return;
            const left = parseFloat(sigOverlay.style.left) || 0;
            const top = parseFloat(sigOverlay.style.top) || 0;
            const width = sigImage.clientWidth;
            const height = sigImage.clientHeight;
            const scaleFactor = 1.5;
            const pdfX = left / scaleFactor;
            const pdfY = top / scaleFactor;
            const pdfW = width / scaleFactor;
            const pdfH = height / scaleFactor;
            
            loading.classList.remove('hidden');
            const formData = new FormData();
            formData.append('file_id', fileId);
            formData.append('page', pageNum - 1);
            formData.append('x', pdfX);
            formData.append('y', pdfY);
            formData.append('w', pdfW);
            formData.append('h', pdfH);
            formData.append('signature', signatureFile);
            
            try {
                const response = await fetch('/add_signature', { method: 'POST', body: formData });
                const data = await response.json();
                if (data.success) {
                    await loadPdf(); 
                    cancelSignature(); 
                } else {
                    alert("Failed to add signature: " + data.error);
                }
            } catch (err) {
                console.error("Signature Error:", err);
                alert("Error connecting to server: " + err.message);
            } finally {
                loading.classList.add('hidden');
            }
        }

        async function extractText() {
            if (!fileId) return;
            if (!APPROVED_USERS.includes(myName)) {
                if(confirm("🔒 Access Restricted\n\nThis premium feature requires admin approval.\n\nClick 'OK' to instantly send a notification request to the admin (Mazharul).")) {
                    fetch('/request_access', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ email: myName })
                    });
                    alert("✅ Your request has been sent! The admin will be notified directly in their portal.");
                }
                return;
            }
            loading.classList.remove('hidden');
            try {
                const response = await fetch('/extract_text', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ file_id: fileId, page: pageNum - 1 }) 
                });
                const data = await response.json();
                if (data.success) {
                    showTextModal(data.text);
                } else {
                    alert("Error extracting text: " + data.error);
                }
            } catch (err) {
                alert("Failed to extract text. Make sure backend is running.");
            } finally {
                loading.classList.add('hidden');
            }
        }

        function showTextModal(text) {
            document.getElementById('extracted-text-area').value = text;
            document.getElementById('text-modal').classList.remove('hidden');
        }

        function closeTextModal() {
            document.getElementById('text-modal').classList.add('hidden');
        }

        function copyExtractedText() {
            const textArea = document.getElementById('extracted-text-area');
            const start = textArea.selectionStart;
            const end = textArea.selectionEnd;
            if (start !== end) {
                textArea.setSelectionRange(start, end);
                document.execCommand('copy');
                alert('Selected text copied to clipboard!');
            } else {
                textArea.select();
                document.execCommand('copy');
                alert('All text copied to clipboard!');
            }
        }

        async function replaceText() {
            if (!fileId) return;
            if (!APPROVED_USERS.includes(myName)) {
                if(confirm("🔒 Access Restricted\n\nThis premium feature requires admin approval.\n\nClick 'OK' to instantly send a notification request to the admin (Mazharul).")) {
                    fetch('/request_access', {
                        method: 'POST',
                        headers: { 'Content-Type': 'application/json' },
                        body: JSON.stringify({ email: myName })
                    });
                    alert("✅ Your request has been sent! The admin will be notified directly in their portal.");
                }
                return;
            }

            const findText = document.getElementById('find-text').value;
            const replaceTextVal = document.getElementById('replace-text').value;
            const occurrence = document.getElementById('occurrence-index').value;
            
            if (!findText) {
                alert("Please enter a word to find.");
                return;
            }

            loading.classList.remove('hidden');
            try {
                const response = await fetch('/replace_text', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ 
                        file_id: fileId, 
                        page: pageNum - 1, 
                        find_text: findText, 
                        replace_text: replaceTextVal,
                        occurrence: occurrence ? parseInt(occurrence, 10) : null
                    }) 
                });
                const data = await response.json();
                if (data.success) {
                    if (data.matches === 0) {
                        alert("Text not found on this page.\n\nTips:\n1. Search for a smaller part of the word (e.g., 'চুক্তি' instead of 'চুক্তিপত্র').\n2. The word might be an image, not text. Try the 'Convert to Word' tool first.");
                    } else {
                        document.getElementById('find-text').value = '';
                        document.getElementById('replace-text').value = '';
                        document.getElementById('occurrence-index').value = '';
                        await loadPdf(); 
                    }
                } else {
                    alert("Error: " + data.error);
                }
            } finally {
                loading.classList.add('hidden');
            }
        }

        // UPDATED: Now uses Save File Picker Dialog
        async function downloadPdf() {
            if (!fileId) return;
            const defaultName = docInfo.textContent !== "No document loaded" ? docInfo.textContent : "edited_document.pdf";
            await saveFileToDisk(`/download/${fileId}`, defaultName, 'application/pdf', '.pdf', 'PDF Document');
        }

        async function downloadNativeFolder(extractId, filesEncoded) {
            if (!window.showDirectoryPicker) {
                alert("Your web browser does not support direct folder saving. Please use Google Chrome or Microsoft Edge, or download the files individually.");
                return;
            }
            const files = JSON.parse(decodeURIComponent(filesEncoded));
            try {
                const dirHandle = await window.showDirectoryPicker({ mode: 'readwrite' });
                loading.classList.remove('hidden');
                for (const file of files) {
                    const pathParts = file.split('/');
                    const filename = pathParts.pop();
                    let currentHandle = dirHandle;
                    for (const part of pathParts) {
                        if (part) {
                            currentHandle = await currentHandle.getDirectoryHandle(part, { create: true });
                        }
                    }
                    const response = await fetch(`/download_extracted/${extractId}/${encodeURIComponent(file)}`);
                    if (!response.ok) continue;
                    const blob = await response.blob();
                    const fileHandle = await currentHandle.getFileHandle(filename, { create: true });
                    const writable = await fileHandle.createWritable();
                    await writable.write(blob);
                    await writable.close();
                }
                alert("Success! All unzipped files have been saved directly to your selected folder.");
            } catch (err) {
                console.error(err);
                if (err.name !== 'AbortError') {
                    alert("Error saving folder. Make sure you grant the browser permission to edit files in that location.");
                }
            } finally {
                loading.classList.add('hidden');
            }
        }
    </script>
</body>
</html>
"""

# -------------------------------------------------------------------
# BACKEND: Flask Routes & Logic
# -------------------------------------------------------------------

@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    file_id = str(uuid.uuid4())
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    file.save(filepath)
    
    doc = fitz.open(filepath)
    pages = len(doc)
    doc.close()
    
    return jsonify({'file_id': file_id, 'pages': pages})

@app.route('/document/<file_id>')
def serve_document(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, mimetype='application/pdf')

# =======================================================================
# NEW: COMPRESS PDF BACKEND ROUTE (WITH AUTO 1MB TARGET)
# =======================================================================
@app.route('/compress_pdf', methods=['POST'])
def compress_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    mode = request.form.get('mode', 'auto_1mb')
    
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Invalid file format. Please upload a PDF.'}), 400
        
    file_id = str(uuid.uuid4())
    orig_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}_orig.pdf")
    comp_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    
    file.save(orig_filepath)
    
    try:
        doc = fitz.open(orig_filepath)
        num_pages = len(doc)
        
        # Handle the Auto 1MB, Auto 512KB or DPI logic
        if mode == 'auto_1mb':
            # Target roughly 950 KB to safely land just under 1 MB and maximize quality
            kb_per_page = 950 / num_pages if num_pages > 0 else 950
            
            # Dynamically choose the best quality that fits the 1MB budget
            if kb_per_page >= 400:
                dpi = 200 # Excellent quality for 1-2 pages
                jpg_qual = 85
            elif kb_per_page >= 200:
                dpi = 150 # High quality for 3-4 pages
                jpg_qual = 80
            elif kb_per_page >= 100:
                dpi = 100 # Medium quality for 5-9 pages
                jpg_qual = 70
            else:
                dpi = 72  # Web quality for 10+ pages
                jpg_qual = 60
                
        elif mode == 'auto_512kb':
            # Target roughly 480 KB to ensure it safely lands under 512 KB
            kb_per_page = 480 / num_pages if num_pages > 0 else 480
            
            # Dynamically choose the best quality that fits the 512KB budget
            if kb_per_page >= 200:
                dpi = 150 # Good quality for 1-2 page documents
            elif kb_per_page >= 100:
                dpi = 100 # Medium quality for 3-4 page documents
            else:
                dpi = 72  # Web quality for 5+ page documents
            jpg_qual = 75 # Slightly better color preservation for this mode
        else:
            # Convert string mode directly to target DPI quality
            try:
                dpi = int(mode)
            except ValueError:
                dpi = 100 # Fallback
            jpg_qual = 70
            
        new_doc = fitz.open()
        for page in doc:
            # Rasterize the page at the chosen DPI quality
            pix = page.get_pixmap(dpi=dpi) 
            # Use a strong JPEG compression quality
            img_data = pix.tobytes("jpeg", jpg_quality=jpg_qual)
            
            # Rebuild the page from the compressed image
            img_doc = fitz.open()
            img_page = img_doc.new_page(width=page.rect.width, height=page.rect.height)
            img_page.insert_image(img_page.rect, stream=img_data)
            
            new_doc.insert_pdf(img_doc)
            img_doc.close()
            
        new_doc.save(comp_filepath, garbage=4, deflate=True)
        new_doc.close()
        doc.close()
        
        # Calculate final size in KB to show the user
        final_size_kb = os.path.getsize(comp_filepath) / 1024
        
        # Clean up the original large file
        if os.path.exists(orig_filepath):
            os.remove(orig_filepath)
            
        return jsonify({'file_id': file_id, 'size_kb': final_size_kb})
    except Exception as e:
        print(f"Error compressing PDF: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/download_compressed/<file_id>')
def download_compressed(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, as_attachment=True, download_name="compressed_document.pdf")

@app.route('/convert_excel', methods=['POST'])
def convert_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    if not file.filename.endswith(('.xls', '.xlsx')):
        return jsonify({'error': 'Invalid format.'}), 400
        
    file_id = str(uuid.uuid4())
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    
    try:
        df = pd.read_excel(file).fillna("")
        pdf = FPDF(orientation='L')
        pdf.add_page()
        pdf.set_font("helvetica", style="B", size=14)
        pdf.cell(0, 10, txt=file.filename, ln=1, align='C')
        pdf.ln(5)
        
        num_cols = len(df.columns)
        epw = pdf.epw 
        col_width = epw / num_cols if num_cols > 0 else 20
        
        pdf.set_font("helvetica", style="B", size=8)
        for col in df.columns:
            header_text = str(col)
            while pdf.get_string_width(header_text + "...") > col_width - 2 and len(header_text) > 0:
                header_text = header_text[:-1]
            if len(header_text) < len(str(col)): header_text += "..."
            pdf.cell(col_width, 10, txt=header_text, border=1) 
        pdf.ln()
        
        pdf.set_font("helvetica", size=7)
        for index, row in df.iterrows():
            for item in row:
                item_text = str(item)
                while pdf.get_string_width(item_text + "...") > col_width - 2 and len(item_text) > 0:
                    item_text = item_text[:-1]
                if len(item_text) < len(str(item)): item_text += "..."
                pdf.cell(col_width, 8, txt=item_text, border=1) 
            pdf.ln()
            
        pdf.output(pdf_filepath)
        return jsonify({'file_id': file_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/merge', methods=['POST'])
def merge_pdfs():
    files = request.files.getlist('files[]')
    if len(files) < 2:
        return jsonify({'error': 'Please upload at least 2 files to merge.'}), 400
        
    merged_doc = fitz.open() 
    file_id = str(uuid.uuid4())
    merged_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    
    opened_docs = []
    temp_paths = []
    
    try:
        for file in files:
            temp_id = str(uuid.uuid4())
            temp_path = os.path.join(UPLOAD_FOLDER, f"{temp_id}.pdf")
            file.save(temp_path)
            temp_paths.append(temp_path)
            
            src_doc = fitz.open(temp_path)
            opened_docs.append(src_doc)
            merged_doc.insert_pdf(src_doc)
            
        merged_doc.save(merged_filepath)
    except Exception as e:
        return jsonify({'error': 'Failed to merge documents.'}), 500
    finally:
        merged_doc.close()
        for doc in opened_docs:
            doc.close()
        for path in temp_paths:
            if os.path.exists(path):
                os.remove(path)
    return jsonify({'file_id': file_id})

@app.route('/convert_pdf_to_excel', methods=['POST'])
def convert_pdf_to_excel():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    file_id = str(uuid.uuid4())
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    excel_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")
    file.save(pdf_filepath)
    
    try:
        wb = openpyxl.Workbook()
        # Remove the default sheet so we can cleanly name our new ones
        default_sheet = wb.active
        wb.remove(default_sheet)
        
        tables_found = False
        
        thin_border = Border(left=Side(style='thin'), 
                             right=Side(style='thin'), 
                             top=Side(style='thin'), 
                             bottom=Side(style='thin'))
        # Align to top-left for a cleaner document look
        wrap_alignment = Alignment(wrap_text=True, vertical='top', horizontal='left')

        with pdfplumber.open(pdf_filepath) as pdf:
            for page_num, page in enumerate(pdf.pages, 1):
                # We try to extract tables based on visible lines first
                tables = page.extract_tables({
                    "vertical_strategy": "lines", 
                    "horizontal_strategy": "lines",
                })
                
                # If no line-based tables exist, guess based on text alignment
                if not tables:
                    tables = page.extract_tables({
                        "vertical_strategy": "text", 
                        "horizontal_strategy": "text",
                        "snap_tolerance": 4, # Slightly higher tolerance to group messy columns
                        "join_tolerance": 4
                    })

                if tables:
                    tables_found = True
                    ws = wb.create_sheet(title=f"Page {page_num}")
                    
                    # HIDE THE DEFAULT EXCEL GRIDLINES to make it look like a white PDF page
                    ws.sheet_view.showGridLines = False 
                    
                    current_row = 1
                    
                    for table in tables:
                        cleaned_table = [[cell if cell is not None else '' for cell in row] for row in table]
                        # Filter completely empty rows
                        cleaned_table = [row for row in cleaned_table if any(str(cell).strip() for cell in row)]
                        
                        if cleaned_table:
                            for row_idx, row_data in enumerate(cleaned_table):
                                for col_idx, cell_value in enumerate(row_data, start=1):
                                    cell = ws.cell(row=current_row + row_idx, column=col_idx, value=cell_value)
                                    cell.alignment = wrap_alignment
                                    
                                    # ONLY apply the border if the cell actually has text.
                                    # This prevents drawing dozens of tiny empty boxes caused by PDF formatting.
                                    if str(cell_value).strip():
                                        cell.border = thin_border
                                    
                            # Add 2 blank rows between different tables on the same page
                            current_row += len(cleaned_table) + 2 
                            
                    # Auto-size columns for this specific page
                    for col in ws.columns:
                        max_length = 0
                        column_letter = col[0].column_letter 
                        for cell in col:
                            if cell.value:
                                try:
                                    lines = str(cell.value).split('\n')
                                    for line in lines:
                                        if len(line) > max_length:
                                            max_length = len(line)
                                except:
                                    pass
                        
                        adjusted_width = (max_length + 2)
                        # Cap the max width so paragraphs force line breaks instead of stretching infinitely
                        ws.column_dimensions[column_letter].width = min(adjusted_width, 50)

        if not tables_found:
            # Revert to a blank workbook if we fail, rather than crashing
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.cell(row=1, column=1, value="No structured tables found on this PDF.")
            wb.save(excel_filepath)
            return jsonify({'file_id': file_id})

        wb.save(excel_filepath)
        return jsonify({'file_id': file_id})
        
    except Exception as e:
        print(f"Error converting PDF to Excel: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/download_excel/<file_id>')
def download_excel(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.xlsx")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, as_attachment=True, download_name="extracted_tables.xlsx")


@app.route('/convert_word_to_pdf', methods=['POST'])
def convert_word_to_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    file_id = str(uuid.uuid4())
    word_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    file.save(word_filepath)
    
    try:
        # 1. Initialize COM for Windows multi-threading (Required for Waitress server)
        pythoncom.CoInitialize()
        
        # 2. Convert losslessly using native MS Word engine
        docx2pdf_convert(word_filepath, pdf_filepath)
        
        return jsonify({'file_id': file_id})
    except Exception as e:
        print(f"Word to PDF Error: {e}")
        return jsonify({'error': "Conversion failed. Please ensure Microsoft Word is installed on the server hosting this portal."}), 500

@app.route('/convert_pdf_to_word', methods=['POST'])
def convert_pdf_to_word():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    file_id = str(uuid.uuid4())
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    word_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    file.save(pdf_filepath)
    
    try:
        cv = Converter(pdf_filepath)
        cv.convert(word_filepath, start=0, end=None)
        cv.close()
        return jsonify({'file_id': file_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download_word/<file_id>')
def download_word(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, as_attachment=True, download_name="converted_document.docx")


@app.route('/convert_pdf_to_jpg', methods=['POST'])
def convert_pdf_to_jpg():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
    file = request.files['file']
    file_id = str(uuid.uuid4())
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    jpg_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.jpg")
    file.save(pdf_filepath)
    
    try:
        doc = fitz.open(pdf_filepath)
        # Convert only the first page to a direct JPG to avoid ZIPs
        if len(doc) > 0:
            page = doc[0]
            pix = page.get_pixmap(dpi=150)
            pix.save(jpg_filepath)
        doc.close()
        return jsonify({'file_id': file_id})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/download_jpg/<file_id>')
def download_jpg(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.jpg")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, mimetype='image/jpeg', as_attachment=True, download_name="converted_image.jpg")

# =======================================================================
# UPDATED: ROUTE FOR MULTIPLE JPG TO SINGLE PDF CONVERSION
# =======================================================================
@app.route('/convert_jpg_to_pdf', methods=['POST'])
def convert_jpg_to_pdf():
    files = request.files.getlist('files[]')
    if not files or len(files) == 0:
        return jsonify({'error': 'No files uploaded'}), 400
        
    file_id = str(uuid.uuid4())
    pdf_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    
    images = []
    temp_filepaths = []
    
    try:
        for file in files:
            if not file.filename.lower().endswith(('.jpg', '.jpeg', '.png')):
                continue
                
            ext = os.path.splitext(file.filename)[1]
            temp_id = str(uuid.uuid4())
            img_filepath = os.path.join(UPLOAD_FOLDER, f"temp_{temp_id}{ext}")
            file.save(img_filepath)
            temp_filepaths.append(img_filepath)
            
            # Open image using PIL
            img = Image.open(img_filepath)
            
            # Convert to RGB to ensure compatibility (PDFs require RGB, not RGBA/P)
            if img.mode != "RGB":
                img = img.convert("RGB")
                
            images.append(img)
            
        if not images:
            return jsonify({'error': 'No valid image files provided.'}), 400

        # Save all images directly into a single multi-page PDF
        images[0].save(
            pdf_filepath, "PDF", resolution=100.0, save_all=True, append_images=images[1:]
        )
        
        # Close all images safely to prevent Windows file locking
        for img in images:
            img.close()
            
        # Clean up the temporary uploaded image files
        for path in temp_filepaths:
            if os.path.exists(path):
                os.remove(path)
                
        return jsonify({'file_id': file_id})
    except Exception as e:
        print(f"Error converting JPG to PDF: {e}")
        # Clean up in case of failure
        for img in images:
            try: img.close()
            except: pass
        for path in temp_filepaths:
            if os.path.exists(path):
                os.remove(path)
        return jsonify({'error': str(e)}), 500

@app.route('/unzip_file', methods=['POST'])
def unzip_file():
    """Unzips an uploaded file and returns the list of extracted paths."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    if not file.filename.lower().endswith('.zip'):
        return jsonify({'error': 'Invalid file format. Please upload a ZIP file.'}), 400
        
    extract_id = str(uuid.uuid4())
    extract_dir = os.path.join(UPLOAD_FOLDER, extract_id)
    os.makedirs(extract_dir, exist_ok=True)
    
    zip_filepath = os.path.join(UPLOAD_FOLDER, f"{extract_id}.zip")
    file.save(zip_filepath)
    
    extracted_files = []
    try:
        with zipfile.ZipFile(zip_filepath, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
            
            # Walk through extracted directory and collect relative file paths
            for root, dirs, files in os.walk(extract_dir):
                for filename in files:
                    full_path = os.path.join(root, filename)
                    rel_path = os.path.relpath(full_path, extract_dir)
                    # Replace backslashes for cross-platform URL compatibility
                    extracted_files.append(rel_path.replace('\\', '/'))
        
        # Clean up the original zip file to save space
        os.remove(zip_filepath)
        return jsonify({'extract_id': extract_id, 'files': extracted_files})
    except Exception as e:
        print(f"Error unzipping file: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/download_extracted/<extract_id>/<path:filename>')
def download_extracted(extract_id, filename):
    """Allows downloading of individual files extracted from a ZIP."""
    # Ensure the path is safe from directory traversal attacks
    safe_path = os.path.normpath(filename)
    if safe_path.startswith('..') or os.path.isabs(safe_path):
        return "Invalid path", 400
        
    filepath = os.path.join(UPLOAD_FOLDER, extract_id, safe_path)
    if not os.path.exists(filepath):
        return "File not found", 404
        
    # Extract just the filename for the download attribute
    download_name = os.path.basename(safe_path)
    return send_file(filepath, as_attachment=True, download_name=download_name)

@app.route('/download_folder/<extract_id>/<filename>')
def download_folder(extract_id, filename):
    """Packages the extracted folder back into a ZIP for bulk downloading."""
    extract_dir = os.path.join(UPLOAD_FOLDER, extract_id)
    if not os.path.exists(extract_dir):
        return "Folder not found", 404
        
    zip_path = os.path.join(UPLOAD_FOLDER, f"{extract_id}_download.zip")
    
    # Create a zip archive of the directory using the standard shutil library
    shutil.make_archive(zip_path.replace('.zip', ''), 'zip', extract_dir)
    
    # Prepend 'extracted_' to the original filename to avoid confusion
    safe_name = "extracted_" + filename if filename.endswith('.zip') else "extracted_folder.zip"
    
    return send_file(zip_path, as_attachment=True, download_name=safe_name)

# =======================================================================
# NEW UPDATED ROUTE: JPG to Editable Word Document via OCR
# =======================================================================
@app.route('/convert_jpg_to_word', methods=['POST'])
def convert_jpg_to_word():
    """Extracts text from an image using OCR and places it into a Word doc."""
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    lang = request.form.get('lang', 'eng') # Safely get language selection

    if not file.filename.lower().endswith(('.jpg', '.jpeg', '.png')):
        return jsonify({'error': 'Invalid file format. Please upload an image.'}), 400
        
    file_id = str(uuid.uuid4())
    ext = os.path.splitext(file.filename)[1]
    img_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}{ext}")
    word_filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.docx")
    
    file.save(img_filepath)
    
    try:
        # 1. Use Tesseract to read the image and extract the string of text
        img = Image.open(img_filepath)
        extracted_text = pytesseract.image_to_string(img, lang=lang) # Apply language
        
        # 2. Create a new Word document
        doc = docx.Document()
        
        # Middle-ground solution: Split by double-newlines (paragraphs)
        # and replace single newlines with spaces to keep sentences together and easily editable.
        if extracted_text.strip():
            paragraphs = extracted_text.split('\n\n')
            for p in paragraphs:
                # Merge broken lines within the same paragraph block
                cleaned_para = p.replace('\n', ' ').strip()
                if cleaned_para:
                    doc.add_paragraph(cleaned_para)
        else:
            doc.add_paragraph("[No readable text was detected in the image]")
            
        doc.save(word_filepath)
        
        return jsonify({'file_id': file_id})
    except Exception as e:
        print(f"Error converting JPG to Word via OCR: {e}")
        return jsonify({'error': str(e)}), 500

# =======================================================================
# NEW: CREATE SECURE / READ-ONLY PDF ROUTE
# =======================================================================
@app.route('/create_secure_pdf', methods=['POST'])
def create_secure_pdf():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400
        
    file = request.files['file']
    filename = file.filename.lower()
    file_id = str(uuid.uuid4())
    temp_ext = os.path.splitext(filename)[1]
    
    input_filepath = os.path.join(UPLOAD_FOLDER, f"input_{file_id}{temp_ext}")
    unsecured_pdf_path = os.path.join(UPLOAD_FOLDER, f"unsecured_{file_id}.pdf")
    locked_pdf_path = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    
    file.save(input_filepath)
    
    try:
        # 1. Standardize everything to a PDF first
        if filename.endswith(('.doc', '.docx')):
            pythoncom.CoInitialize()
            docx2pdf_convert(input_filepath, unsecured_pdf_path)
        elif filename.endswith(('.jpg', '.jpeg', '.png')):
            img = Image.open(input_filepath)
            if img.mode != "RGB":
                img = img.convert("RGB")
            img.save(unsecured_pdf_path, "PDF", resolution=100.0)
            img.close()
        elif filename.endswith('.pdf'):
            shutil.copy(input_filepath, unsecured_pdf_path)
        else:
            return jsonify({'error': 'Unsupported file format.'}), 400
            
        # 2. Lock the PDF with PyMuPDF
        doc = fitz.open(unsecured_pdf_path)
        
        # Allow PRINT and ACCESSIBILITY, but restrict copying, modifying, and assembling
        perms = fitz.PDF_PERM_PRINT | fitz.PDF_PERM_ACCESSIBILITY
        
        # Using a random owner password makes it impossible for standard tools to unlock editing
        random_owner_pw = str(uuid.uuid4())
        
        doc.save(locked_pdf_path, 
                 encryption=fitz.PDF_ENCRYPT_AES_256, 
                 owner_pw=random_owner_pw, 
                 user_pw="",  # Blank user password = anyone can open and read
                 permissions=perms)
        doc.close()
        
        # Cleanup temporary processing files
        if os.path.exists(input_filepath): os.remove(input_filepath)
        if os.path.exists(unsecured_pdf_path): os.remove(unsecured_pdf_path)
        
        return jsonify({'file_id': file_id})
        
    except Exception as e:
        print(f"Error securing PDF: {e}")
        if os.path.exists(input_filepath): os.remove(input_filepath)
        if os.path.exists(unsecured_pdf_path): os.remove(unsecured_pdf_path)
        return jsonify({'error': str(e)}), 500

@app.route('/rotate_page', methods=['POST'])
def rotate_page():
    data = request.json
    file_id = data.get('file_id')
    page_num = data.get('page') 
    rotation_amount = data.get('rotation', 90)
    
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
        
    doc = fitz.open(filepath)
    if 0 <= page_num < len(doc):
        page = doc[page_num]
        current_rotation = page.rotation
        page.set_rotation((current_rotation + rotation_amount) % 360)
        
        temp_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_temp.pdf")
        doc.save(temp_path)
        doc.close()
        os.replace(temp_path, filepath)
        return jsonify({'success': True})
        
    doc.close()
    return jsonify({'error': 'Invalid page number'}), 400


@app.route('/delete_page', methods=['POST'])
def delete_page():
    data = request.json
    file_id = data.get('file_id')
    page_num = data.get('page')
    
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
        
    doc = fitz.open(filepath)
    if 0 <= page_num < len(doc):
        doc.delete_page(page_num)
        temp_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_temp.pdf")
        doc.save(temp_path)
        doc.close()
        os.replace(temp_path, filepath) 
        
        new_doc = fitz.open(filepath)
        new_len = len(new_doc)
        new_doc.close()
        return jsonify({'success': True, 'total_pages': new_len})
        
    doc.close()
    return jsonify({'error': 'Invalid page number'}), 400

# =======================================================================
# NEW: INSERT PDF IN THE MIDDLE OF ANOTHER PDF
# =======================================================================
@app.route('/insert_pdf', methods=['POST'])
def insert_pdf():
    file_id = request.form.get('file_id')
    insert_at = int(request.form.get('page', -1))
    
    if not file_id:
        return jsonify({'error': 'Missing file ID'}), 400
        
    if 'file' not in request.files:
        return jsonify({'error': 'No PDF file uploaded to insert'}), 400
        
    insert_file = request.files['file']
    if not insert_file.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a valid PDF file to insert.'}), 400
        
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return jsonify({'error': 'Original PDF not found'}), 404
        
    # Save the new PDF to insert temporarily
    temp_insert_id = str(uuid.uuid4())
    temp_insert_path = os.path.join(UPLOAD_FOLDER, f"insert_{temp_insert_id}.pdf")
    insert_file.save(temp_insert_path)
    
    try:
        main_doc = fitz.open(filepath)
        insert_doc = fitz.open(temp_insert_path)
        inserted_count = len(insert_doc)
        
        # Insert the secondary document into the main document exactly at the target index
        main_doc.insert_pdf(insert_doc, start_at=insert_at)
        
        temp_main_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_temp.pdf")
        main_doc.save(temp_main_path)
        
        main_doc.close()
        insert_doc.close()
        
        # Overwrite the original file with the newly combined one
        os.replace(temp_main_path, filepath)
        
        # Clean up the temporary uploaded file
        if os.path.exists(temp_insert_path):
            os.remove(temp_insert_path)
            
        return jsonify({'success': True, 'inserted_pages': inserted_count})
        
    except Exception as e:
        print(f"Error inserting PDF: {e}")
        if 'main_doc' in locals(): main_doc.close()
        if 'insert_doc' in locals(): insert_doc.close()
        if os.path.exists(temp_insert_path): os.remove(temp_insert_path)
        return jsonify({'error': str(e)}), 500

# NEW: Backend route to pull all text from the current page
@app.route('/extract_text', methods=['POST'])
def extract_text():
    data = request.json
    file_id = data.get('file_id')
    page_num = data.get('page')
    
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
        
    try:
        doc = fitz.open(filepath)
        if 0 <= page_num < len(doc):
            page = doc[page_num]
            text = page.get_text("text")
            
            # --- AUTO-DETECT BIJOY (ANSI BANGLA) ---
            is_bijoy = False
            
            # Check 1: Inspect font names embedded inside the PDF page
            fonts = page.get_fonts()
            for font in fonts:
                font_name = font[3].lower() if len(font) > 3 else ""
                if any(x in font_name for x in ['sutonny', 'bijoy', 'siyam', 'boishakhi', 'kalpurush']):
                    is_bijoy = True
                    break
                    
            # Check 2: Heuristic detection of extended ASCII symbols heavily used in Bijoy layout
            bijoy_markers = ['‡', 'ÿ', 'Ø', '¤', '¨', 'œ', 'Š', 'µ', '¶', '·', '¸', '¹', 'º', '»', '˜', '™', 'š', '›', 'ž', 'Ÿ', '¢', '£', '¥', '¦', '§', '©', 'ª', '«', '¬', '®', '¯', '°', '±', '²', '³', '´']
            if not is_bijoy and sum(1 for c in text if c in bijoy_markers) > 2:
                is_bijoy = True

            # If detected, silently and automatically convert to perfect Unicode!
            if is_bijoy:
                try:
                    from bijoy2unicode import converter
                    un = converter.Unicode()
                    fixed_text = un.convertBijoyToUnicode(text)
                    
                    # Post-processing fixes for known missing glyphs
                    corrections = {
                        'ÿ': 'ক্ষ',
                        'ন্‌ত্ম': 'ন্ত',
                        'ফ্যাট': 'ফ্ল্যাট'
                    }
                    for broken, fixed in corrections.items():
                        fixed_text = fixed_text.replace(broken, fixed)
                        
                    text = fixed_text
                except ImportError:
                    print("bijoy2unicode library missing. Skipping auto-conversion.")
                except Exception as e:
                    print(f"Bijoy conversion error: {e}")

            return jsonify({'success': True, 'text': text})
        return jsonify({'error': 'Invalid page number'}), 400
    except Exception as e:
        return jsonify({'error': str(e)}), 500
    finally:
        if 'doc' in locals():
            doc.close()

@app.route('/replace_text', methods=['POST'])
def replace_text():
    data = request.json
    file_id = data.get('file_id')
    page_num = data.get('page') 
    find_text = data.get('find_text')
    replace_val = data.get('replace_text', '')
    occurrence = data.get('occurrence')
    
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return jsonify({'error': 'File not found'}), 404
        
    import re
    def clean_string(s):
        """Removes invisible characters, punctuation, spaces, and normalizes Bengali digits to English digits for bulletproof matching."""
        if not s: return ""
        s = s.replace('\u200c', '').replace('\u200d', '') # Remove invisible Joiners
        s = re.sub(r'[\s\.\-\,\:\;\(\)\[\]\{\}\_\|\'\"\”\“\‘\’\?\/\\+]+', '', s) # Remove spaces and punctuation
        s = s.translate(str.maketrans('০১২৩৪৫৬৭৮৯', '0123456789')) # Normalize numbers
        return s
        
    doc = fitz.open(filepath)
    if 0 <= page_num < len(doc):
        page = doc[page_num]
        
        # 1. Try standard exact search
        text_instances = page.search_for(find_text)
        
        # 2. Advanced Fuzzy & Deep Translation Search 
        if not text_instances:
            try:
                from bijoy2unicode import converter
                un = converter.Unicode()
            except:
                un = None
                
            corrections = {'ÿ': 'ক্ষ', 'ন্‌ত্ম': 'ন্ত', 'ফ্যাট': 'ফ্ল্যাট', '': ''}
            find_clean = clean_string(find_text)
            is_bangla_search = any(ord(c) > 127 for c in find_text)
            
            if find_clean:
                words = page.get_text("words")
                
                # Pre-process all words independently so a single bad char doesn't fail a whole phrase
                processed_words = []
                for w in words:
                    raw_str = w[4]
                    rect = fitz.Rect(w[:4])
                    uni_str = raw_str
                    
                    if is_bangla_search and un:
                        try:
                            temp = un.convertBijoyToUnicode(raw_str)
                            for broken, fixed in corrections.items():
                                temp = temp.replace(broken, fixed)
                            uni_str = temp
                        except:
                            pass
                            
                    cleaned_str = clean_string(uni_str)
                    if cleaned_str:
                        processed_words.append({
                            "rect": rect,
                            "clean": cleaned_str
                        })
                        
                # Sliding window over the cleaned, translated words
                for i in range(len(processed_words)):
                    current_rect = processed_words[i]["rect"]
                    combined_text = ""
                    
                    for j in range(i, min(i + 15, len(processed_words))):
                        if j > i:
                            current_rect = current_rect.union(processed_words[j]["rect"])
                        combined_text += processed_words[j]["clean"]
                        
                        if find_clean in combined_text:
                            if not any(r.intersects(current_rect) for r in text_instances):
                                text_instances.append(current_rect)
                            break
                            
                        if len(combined_text) > len(find_clean) + 50:
                            break
        
        if text_instances:
            # If an occurrence number was specified, filter the list down to just that one
            if occurrence is not None:
                if 1 <= occurrence <= len(text_instances):
                    text_instances = [text_instances[occurrence - 1]]
                else:
                    doc.close()
                    return jsonify({'error': f'Instance {occurrence} not found. There are only {len(text_instances)} matches for that word.'}), 400

            text_dict = page.get_text("dict")
            replacements = []
            
            for inst in text_instances:
                detected_font_size = 11
                detected_color = (0, 0, 0)
                detected_fontname = "helv" # Default to standard Helvetica
                # Approximate baseline if not found
                baseline_y = inst.y1 - (inst.height * 0.2)
                
                max_overlap = 0
                for block in text_dict.get("blocks", []):
                    if block.get("type") == 0: 
                        for line in block.get("lines", []):
                            for span in line.get("spans", []):
                                span_rect = fitz.Rect(span["bbox"])
                                # Calculate how much the found word overlaps with this text span
                                overlap = span_rect.intersect(inst)
                                overlap_area = overlap.get_area()
                                
                                if overlap_area > max_overlap:
                                    max_overlap = overlap_area
                                    detected_font_size = span["size"]
                                    baseline_y = span["origin"][1]
                                    
                                    # Extract the exact font color
                                    c = span["color"]
                                    detected_color = (((c >> 16) & 255)/255, ((c >> 8) & 255)/255, (c & 255)/255)
                                    
                                    # Detect if the original font is bold
                                    is_bold = False
                                    if "flags" in span and (span["flags"] & 16): # Bit 4 (16) indicates bold
                                        is_bold = True
                                    elif "font" in span and "bold" in span["font"].lower():
                                        is_bold = True
                                        
                                    detected_fontname = "hebo" if is_bold else "helv"
                
                # Store the exact coordinates and styles for the new text
                replacements.append({
                    "point": fitz.Point(inst.x0, baseline_y),
                    "size": detected_font_size,
                    "color": detected_color,
                    "fontname": detected_fontname
                })
                
                # UPDATED: Use completely transparent redaction to delete the old text.
                # cross_out=False ensures no red lines are drawn, and fill=None ensures no boxes are painted.
                # This perfectly preserves complex backgrounds like the Bangladesh watermark seal!
                page.add_redact_annot(inst, fill=None, cross_out=False)
            
            # Apply redactions to permanently remove the old text
            # images=0 prevents erasing background images/watermarks
            # graphics=0 prevents erasing vector graphics and lines
            page.apply_redactions(images=0, graphics=0)
            
            # --- NEW: LOAD BANGLA FONT IF NEEDED ---
            # PDFs natively support only English characters. To write Bangla, we load our custom font file.
            font_path = "bangla.ttf" # Must be in the same folder as app.py
            has_custom_font = False
            
            # Check if the user is trying to write Bengali/Unicode characters
            is_bangla_replacement = any(ord(char) > 127 for char in replace_val)
            
            if is_bangla_replacement:
                if os.path.exists(font_path):
                    try:
                        page.insert_font(fontname="bang", fontfile=font_path)
                        has_custom_font = True
                    except Exception as e:
                        print(f"Could not load Bangla font: {e}")
                else:
                    # STRICT CHECK: Do not fail silently. Warn the user they are missing the font file!
                    doc.close()
                    return jsonify({'error': 'Missing Font! Please download a Bengali font (like Kalpurush.ttf), rename it to "bangla.ttf", and put it in your server folder to write Bangla.'}), 400

            # Insert the new text using the exact matching font weight, size, and color
            if replace_val:
                for rep in replacements:
                    use_font = rep["fontname"]
                    
                    # If the user typed non-English characters (like Bangla) and we have the font file, use it!
                    if has_custom_font and is_bangla_replacement:
                        use_font = "bang"
                        
                    page.insert_text(rep["point"], replace_val, fontname=use_font, fontsize=rep["size"], color=rep["color"])
            
            temp_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_temp.pdf")
            doc.save(temp_path)
            doc.close()
            os.replace(temp_path, filepath)
            return jsonify({'success': True, 'matches': len(text_instances)})
        else:
            doc.close()
            return jsonify({'success': True, 'matches': 0})
            
    doc.close()
    return jsonify({'error': 'Invalid page number'}), 400

# =======================================================================
# NEW: ADD SIGNATURE ROUTE
# =======================================================================
@app.route('/add_signature', methods=['POST'])
def add_signature():
    try:
        file_id = request.form.get('file_id')
        if not file_id:
            return jsonify({'error': 'Missing file ID'}), 400
            
        page_num = int(request.form.get('page', 0))
        x = float(request.form.get('x', 0))
        y = float(request.form.get('y', 0))
        w = float(request.form.get('w', 0))
        h = float(request.form.get('h', 0))
        
        if 'signature' not in request.files:
            return jsonify({'error': 'No signature file uploaded'}), 400
            
        sig_file = request.files['signature']
        filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
        
        if not os.path.exists(filepath):
            return jsonify({'error': 'PDF not found'}), 404
            
        # FIX 1: Use the original file extension so PyMuPDF doesn't crash on format mismatch
        ext = os.path.splitext(sig_file.filename)[1].lower()
        if ext not in ['.png', '.jpg', '.jpeg']:
            ext = '.png'
            
        # Save the uploaded signature image temporarily
        sig_path = os.path.join(UPLOAD_FOLDER, f"sig_{uuid.uuid4()}{ext}")
        sig_file.save(sig_path)
        
        doc = fitz.open(filepath)
        if 0 <= page_num < len(doc):
            page = doc[page_num]
            
            # Read the image size using PIL to maintain aspect ratio
            img = Image.open(sig_path)
            img_w, img_h = img.size
            img.close() 
            
            # Apply dynamic width/height from frontend user, or fallback to default
            if w > 0 and h > 0:
                target_w = w
                target_h = h
            else:
                target_w = 150
                target_h = target_w * (img_h / img_w)
            
            # Create the rectangle where the signature will be stamped
            # fitz.Rect(x0, y0, x1, y1)
            rect = fitz.Rect(x, y, x + target_w, y + target_h)
            
            # Insert the image
            page.insert_image(rect, filename=sig_path)
            
            temp_path = os.path.join(UPLOAD_FOLDER, f"{file_id}_temp.pdf")
            doc.save(temp_path)
            doc.close()
            os.replace(temp_path, filepath)
            
            # Clean up the temporary signature image safely
            if os.path.exists(sig_path):
                os.remove(sig_path)
            return jsonify({'success': True})
            
        doc.close()
        if os.path.exists(sig_path):
            os.remove(sig_path)
        return jsonify({'error': 'Invalid page number'}), 400
        
    except Exception as e:
        print(f"Error adding signature: {e}")
        # Attempt to clean up the file if it crashed during the process
        try:
            if 'sig_path' in locals() and os.path.exists(sig_path):
                os.remove(sig_path)
        except:
            pass
        return jsonify({'error': str(e)}), 500

@app.route('/download/<file_id>')
def download_file(file_id):
    filepath = os.path.join(UPLOAD_FOLDER, f"{file_id}.pdf")
    if not os.path.exists(filepath):
        return "File not found", 404
    return send_file(filepath, mimetype='application/pdf', as_attachment=True, download_name="edited_document.pdf")

# =======================================================================
# NEW: LIVE USER HEARTBEAT & LOGGING ROUTE
# =======================================================================
@app.route('/heartbeat', methods=['POST'])
def heartbeat():
    data = request.json
    session_id = data.get('session_id')
    name = data.get('name')
    
    if session_id and name:
        # If this is a completely new session, write it to our usage log!
        if session_id not in ACTIVE_USERS:
            with open(USAGE_LOG, 'a', encoding='utf-8') as f:
                timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
                ip_address = request.remote_addr
                f.write(f"{timestamp},{name},{session_id},{ip_address}\n")
                
        ACTIVE_USERS[session_id] = {
            'name': name,
            'last_seen': time.time()
        }
        
    # Clean up inactive users (anyone who hasn't pinged in 30 seconds)
    current_time = time.time()
    stale_sessions = [sid for sid, info in ACTIVE_USERS.items() if current_time - info['last_seen'] > 30]
    for sid in stale_sessions:
        del ACTIVE_USERS[sid]
        
    # Get list of unique online names
    online_names = list(set([info['name'] for info in ACTIVE_USERS.values()]))
    
    response_data = {'users': sorted(online_names)}
    
    # NEW: Send the list of approved users to the frontend
    if os.path.exists(APPROVED_USERS_FILE):
        with open(APPROVED_USERS_FILE, 'r', encoding='utf-8') as f:
            response_data['approved_users'] = [line.strip() for line in f.readlines() if line.strip()]
    
    # NEW: If the admin is polling, send them the list of pending access requests!
    if name == 'haque.mazharul@ap.averydennison.com':
        if os.path.exists(REQUESTS_FILE):
            with open(REQUESTS_FILE, 'r', encoding='utf-8') as f:
                pending = [line.strip() for line in f.readlines() if line.strip()]
            response_data['pending_requests'] = pending
            
    return jsonify(response_data)

# =======================================================================
# NEW: IN-APP NOTIFICATION ROUTES
# =======================================================================
@app.route('/register_user', methods=['POST'])
def register_user():
    data = request.json
    email = data.get('email')
    if email:
        with open(REGISTRATION_LOG, 'a', encoding='utf-8') as f:
            timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
            ip_address = request.remote_addr
            f.write(f"{timestamp},{email},{ip_address}\n")
        return jsonify({'success': True})
    return jsonify({'error': 'No email'}), 400

@app.route('/request_access', methods=['POST'])
def request_access():
    email = request.json.get('email')
    if email:
        # Read existing to avoid duplicates
        existing = []
        if os.path.exists(REQUESTS_FILE):
            with open(REQUESTS_FILE, 'r', encoding='utf-8') as f:
                existing = f.read().splitlines()
                
        if email not in existing:
            with open(REQUESTS_FILE, 'a', encoding='utf-8') as f:
                f.write(f"{email}\n")
        return jsonify({'success': True})
    return jsonify({'error': 'No email'}), 400

@app.route('/approve_users', methods=['POST'])
def approve_users():
    # Admin clicked "OK" to approve pending users
    if os.path.exists(REQUESTS_FILE):
        with open(REQUESTS_FILE, 'r', encoding='utf-8') as f:
            pending = [line.strip() for line in f.readlines() if line.strip()]
            
        # Read currently approved to avoid duplicates
        approved = []
        if os.path.exists(APPROVED_USERS_FILE):
            with open(APPROVED_USERS_FILE, 'r', encoding='utf-8') as f:
                approved = [line.strip() for line in f.readlines() if line.strip()]
        
        # Add new users to the approved list
        with open(APPROVED_USERS_FILE, 'a', encoding='utf-8') as f:
            for email in pending:
                if email not in approved:
                    f.write(f"{email}\n")
                    
        # Clear the notification queue
        with open(REQUESTS_FILE, 'w', encoding='utf-8') as f:
            pass
            
    return jsonify({'success': True})

@app.route('/clear_requests', methods=['POST'])
def clear_requests():
    # Admin clicked to clear notifications
    with open(REQUESTS_FILE, 'w', encoding='utf-8') as f:
        pass
    return jsonify({'success': True})

# =======================================================================
# UPDATED: ROUTE TO DOWNLOAD THE LOG FILE (NOW PROTECTED)
# =======================================================================
@app.route('/download_logs')
def download_logs():
    # Check if the provided password matches
    admin_password = "paxar@123" 
    provided_key = request.args.get('key')
    log_type = request.args.get('type', 'usage')
    
    if provided_key != admin_password:
        return "Unauthorized: Incorrect Admin Password", 401
        
    target_file = REGISTRATION_LOG if log_type == 'registration' else USAGE_LOG
    dl_name = "DocHub_Registration_Logs.csv" if log_type == 'registration' else "DocHub_Usage_Logs.csv"
    
    if os.path.exists(target_file):
        return send_file(target_file, as_attachment=True, download_name=dl_name)
    return "No logs generated yet.", 404

if __name__ == '__main__':
    from waitress import serve
    print("🚀 Starting DocHub PDF Portal on Network Port 5000")
    print("🛡️  Running securely via Waitress Production Server...")
    serve(app, host='0.0.0.0', port=5000)