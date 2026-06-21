import {
	IExecuteFunctions,
	INodeExecutionData,
	INodeType,
	INodeTypeDescription,
	NodeOperationError,
} from 'n8n-workflow';
import pdfParse from 'pdf-parse/lib/pdf-parse';
import WordExtractor from 'word-extractor';
import mammoth from 'mammoth';
import * as XLSX from 'xlsx';
import TurndownService from 'turndown';

// Helper to determine page ranges for PDF parsing
export function isPageInRange(pageNum: number, rangeStr: string): boolean {
	if (!rangeStr || rangeStr.trim() === '') return true;
	const parts = rangeStr.split(',').map(p => p.trim());
	for (const part of parts) {
		if (part.includes('-')) {
			const [startStr, endStr] = part.split('-');
			const start = parseInt(startStr, 10);
			const end = parseInt(endStr, 10);
			if (!isNaN(start) && !isNaN(end)) {
				if (pageNum >= start && pageNum <= end) return true;
			}
		} else {
			const single = parseInt(part, 10);
			if (!isNaN(single)) {
				if (pageNum === single) return true;
			}
		}
	}
	return false;
}

// Convert 0 -> A, 1 -> B, 25 -> Z, 26 -> AA etc.
export function getColName(colIdx: number): string {
	let name = '';
	let temp = colIdx;
	while (temp >= 0) {
		name = String.fromCharCode((temp % 26) + 65) + name;
		temp = Math.floor(temp / 26) - 1;
	}
	return name;
}

// Convert sheet to Markdown Table (GitHub Flavored Markdown)
export function sheetToMarkdownTable(sheet: XLSX.WorkSheet, hasHeader: boolean): string {
	const rows = XLSX.utils.sheet_to_json(sheet, { header: 1 }) as any[][];
	if (rows.length === 0) return '';
	let md = '';

	let headers: string[] = [];
	let startRow = 0;

	if (hasHeader) {
		headers = (rows[0] || []).map(h => String(h ?? '').trim());
		startRow = 1;
	} else {
		const maxCols = Math.max(...rows.map(r => r.length), 0);
		for (let c = 0; c < maxCols; c++) {
			headers.push(getColName(c));
		}
		startRow = 0;
	}

	if (headers.length === 0) return '';

	md += '| ' + headers.join(' | ') + ' |\n';
	md += '| ' + headers.map(() => '---').join(' | ') + ' |\n';

	for (let r = startRow; r < rows.length; r++) {
		const row = rows[r] || [];
		const cells = headers.map((_, idx) => {
			const val = row[idx];
			return String(val ?? '')
				.replace(/[\r\n]+/g, '<br>')
				.replace(/\|/g, '\\|')
				.trim();
		});
		md += '| ' + cells.join(' | ') + ' |\n';
	}

	return md;
}

// ──────────────────────────────────────────────────────────────────────────────
// Output Cleanup Helper
// ──────────────────────────────────────────────────────────────────────────────
export interface CleanupOptions {
	removeBase64: boolean;
	removeStyles: boolean;
	removeScripts: boolean;
	collapseWhitespace: boolean;
	removeComments: boolean;
	removeEmptyTags: boolean;
}

export function applyCleanup(text: string, opts: CleanupOptions): string {
	let result = text;

	// 1. Remove base64 encoded data (data URIs and raw base64 blobs ≥ 64 chars)
	if (opts.removeBase64) {
		// Remove data URI attributes: src="data:..." or href="data:..."
		result = result.replace(/(?:src|href|data)="data:[^"]{20,}"/gi, '');
		// Remove bare base64 strings inside tags (e.g. style background-image)
		result = result.replace(/data:[a-z/]+;base64,[A-Za-z0-9+/=]{20,}/gi, '');
		// Remove standalone base64-looking blobs (≥ 64 chars of base64 chars) on their own line
		result = result.replace(/^[A-Za-z0-9+/]{64,}={0,2}$/gm, '');
	}

	// 2. Remove inline style, class, id, width, height, align, bgcolor, color attributes
	if (opts.removeStyles) {
		result = result.replace(/\s+style="[^"]*"/gi, '');
		result = result.replace(/\s+class="[^"]*"/gi, '');
		result = result.replace(/\s+id="[^"]*"/gi, '');
		result = result.replace(/\s+width="[^"]*"/gi, '');
		result = result.replace(/\s+height="[^"]*"/gi, '');
		result = result.replace(/\s+align="[^"]*"/gi, '');
		result = result.replace(/\s+bgcolor="[^"]*"/gi, '');
		result = result.replace(/\s+color="[^"]*"/gi, '');
		result = result.replace(/\s+valign="[^"]*"/gi, '');
		result = result.replace(/\s+cellpadding="[^"]*"/gi, '');
		result = result.replace(/\s+cellspacing="[^"]*"/gi, '');
		result = result.replace(/\s+border="[^"]*"/gi, '');
	}

	// 3. Remove <script> and <style> blocks entirely
	if (opts.removeScripts) {
		result = result.replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi, '');
		result = result.replace(/<style\b[^>]*>[\s\S]*?<\/style>/gi, '');
		result = result.replace(/<noscript\b[^>]*>[\s\S]*?<\/noscript>/gi, '');
	}

	// 4. Remove HTML comments
	if (opts.removeComments) {
		result = result.replace(/<!--[\s\S]*?-->/g, '');
	}

	// 5. Remove empty/useless tags (tags with only whitespace inside)
	if (opts.removeEmptyTags) {
		// Repeat up to 3 times to catch nested empty tags
		for (let pass = 0; pass < 3; pass++) {
			result = result.replace(/<(span|div|p|em|strong|b|i|u|s|a|label|li|td|th)[^>]*>\s*<\/\1>/gi, '');
		}
	}

	// 6. Collapse excess whitespace / blank lines
	if (opts.collapseWhitespace) {
		// Replace 3+ consecutive newlines with 2
		result = result.replace(/\n{3,}/g, '\n\n');
		// Replace lines containing only spaces/tabs with empty lines
		result = result.replace(/^[ \t]+$/gm, '');
		// Trim leading/trailing whitespace per line for text content
		result = result.replace(/[ \t]{2,}/g, ' ');
		result = result.trim();
	}

	return result;
}

export class Doc4AI implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'Doc4AI',
		name: 'doc4ai',
		icon: 'file:doc4ai-icon.png',
		group: ['transform'],
		version: [1],
		subtitle: '={{$parameter["operation"]}}',
		description: 'Convert document files (PDF, Word, Excel, CSV) to AI-friendly text formats or generate files from text.',
		defaults: {
			name: 'Doc4AI',
		},
		usableAsTool: true,
		inputs: ['main'],
		outputs: ['main'],
		properties: [
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				options: [
					{
						name: 'Extract Text from File',
						value: 'extractText',
						description: 'Extract and format document contents as Markdown, HTML, or Plain Text',
						action: 'Extract text from a file',
					},
					{
						name: 'Convert Text to File',
						value: 'convertToFile',
						description: 'Generate binary files (XLSX, CSV, HTML, TXT) from text or JSON data',
						action: 'Convert text to a file',
					},
				],
				default: 'extractText',
			},

			// ==========================================
			// Parameters for: Extract Text
			// ==========================================
			{
				displayName: 'Binary Property',
				name: 'binaryPropertyName',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						operation: ['extractText'],
					},
				},
				description: 'Name of the binary property containing the file to convert',
			},
			{
				displayName: 'Output Format',
				name: 'outputFormat',
				type: 'options',
				options: [
					{ name: 'Markdown', value: 'markdown' },
					{ name: 'HTML', value: 'html' },
					{ name: 'Plain Text', value: 'text' },
				],
				default: 'markdown',
				displayOptions: {
					show: {
						operation: ['extractText'],
					},
				},
				description: 'The target format to output the extracted text in',
			},
			// ==========================================
			// Parameters for: Convert Text to File
			// ==========================================
			{
				displayName: 'File Format',
				name: 'fileFormat',
				type: 'options',
				options: [
					{ name: 'CSV', value: 'csv' },
					{ name: 'Excel (XLSX)', value: 'xlsx' },
					{ name: 'HTML', value: 'html' },
					{ name: 'Plain Text', value: 'txt' },
				],
				default: 'xlsx',
				displayOptions: {
					show: {
						operation: ['convertToFile'],
					},
				},
				description: 'The target file format to output',
			},
			{
				displayName: 'Source Field',
				name: 'sourceKey',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						operation: ['convertToFile'],
					},
				},
				description: 'The JSON field name containing the data to convert (string for text/HTML, array/objects for CSV/XLSX)',
			},
			{
				displayName: 'File Name',
				name: 'fileName',
				type: 'string',
				default: 'document',
				required: true,
				displayOptions: {
					show: {
						operation: ['convertToFile'],
					},
				},
				description: 'The name of the generated file (without extension, extension will be added automatically)',
			},
			{
				displayName: 'Binary Property (Output)',
				name: 'binaryPropertyNameOutput',
				type: 'string',
				default: 'data',
				required: true,
				displayOptions: {
					show: {
						operation: ['convertToFile'],
					},
				},
				description: 'Name of the binary property in the output item where the file will be stored',
			},
			{
				displayName: 'Additional Fields',
				name: 'additionalFields',
				type: 'collection',
				placeholder: 'Add Options',
				default: {},
				options: [
					// Extract Text Options
					{
						displayName: 'Include Character Counts',
						name: 'includeCharCounts',
						type: 'boolean',
						displayOptions: {
							show: {
								'/operation': ['extractText'],
							},
						},
						default: false,
						description: 'Whether to return original character count, converted character count, and byte size',
					},
					{
						displayName: 'Maximum File Size (MB)',
						name: 'maxFileSize',
						type: 'number',
						displayOptions: {
							show: {
								'/operation': ['extractText'],
							},
						},
						default: 30,
						description: 'The maximum allowed size of the file in MB. Files larger than this will be rejected.',
					},
					{
						displayName: 'PDF Page Range',
						name: 'pdfPageRange',
						type: 'string',
						displayOptions: {
							show: {
								'/operation': ['extractText'],
							},
						},
						default: '',
						placeholder: 'e.g., 1-5, 8, 11-13',
						description: 'Specify which pages to extract for PDFs. Leave blank to extract all pages.',
					},
					{
						displayName: 'Excel Sheet Name',
						name: 'excelSheetName',
						type: 'string',
						displayOptions: {
							show: {
								'/operation': ['extractText'],
							},
						},
						default: '',
						placeholder: 'Sheet1',
						description: 'Specify which sheet to extract for Excel files. Leave blank to extract all sheets.',
					},
					{
						displayName: 'Excel Has Header Row',
						name: 'excelHeaderRow',
						type: 'boolean',
						displayOptions: {
							show: {
								'/operation': ['extractText'],
							},
						},
						default: true,
						description: 'Whether the first row of Excel sheets should be treated as headers',
					},
					// Convert to File Options
					{
						displayName: 'Excel Sheet Name',
						name: 'excelSheetNameOutput',
						type: 'string',
						displayOptions: {
							show: {
								'/operation': ['convertToFile'],
								'/fileFormat': ['xlsx'],
							},
						},
						default: 'Sheet1',
						description: 'Sheet name to create in Excel workbook',
					},
					{
						displayName: 'CSV Delimiter',
						name: 'csvDelimiter',
						type: 'string',
						displayOptions: {
							show: {
								'/operation': ['convertToFile'],
								'/fileFormat': ['csv'],
							},
						},
						default: ',',
						description: 'The character used to separate values in CSV outputs',
					},
				],
			},
			// ==========================================
			// Output Cleanup Options (Extract Text only)
			// ==========================================
			{
				displayName: 'Output Cleanup',
				name: 'cleanupOptions',
				type: 'collection',
				placeholder: 'Add Cleanup Option',
				default: {},
				displayOptions: {
					show: {
						operation: ['extractText'],
					},
				},
				description: 'Options to clean and simplify the extracted text for AI processing. All options are enabled by default to keep only essential text and structure.',
				options: [
					{
						displayName: 'Remove Base64 Data',
						name: 'removeBase64',
						type: 'boolean',
						default: true,
						description: 'Whether to remove base64-encoded images and embedded binary data (data URIs). Greatly reduces output size.',
					},
					{
						displayName: 'Remove Inline Styles & Attributes',
						name: 'removeStyles',
						type: 'boolean',
						default: true,
						description: 'Whether to strip style="", class="", id="", width="", height="" and other presentational HTML attributes that are unnecessary for AI understanding.',
					},
					{
						displayName: 'Remove Scripts & Style Blocks',
						name: 'removeScripts',
						type: 'boolean',
						default: true,
						description: 'Whether to remove <script>, <style>, and <noscript> blocks entirely from HTML output.',
					},
					{
						displayName: 'Remove HTML Comments',
						name: 'removeComments',
						type: 'boolean',
						default: true,
						description: 'Whether to remove HTML comments (<!-- ... -->) from the output.',
					},
					{
						displayName: 'Remove Empty Tags',
						name: 'removeEmptyTags',
						type: 'boolean',
						default: true,
						description: 'Whether to remove empty HTML tags that contain no content (e.g. <span></span>, <p></p>).',
					},
					{
						displayName: 'Collapse Excess Whitespace',
						name: 'collapseWhitespace',
						type: 'boolean',
						default: true,
						description: 'Whether to collapse 3+ consecutive blank lines into 2, remove trailing spaces per line, and trim the result.',
					},
				],
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const turndownService = new TurndownService();

		for (let i = 0; i < items.length; i++) {
			try {
				const operation = this.getNodeParameter('operation', i) as string;

				if (operation === 'extractText') {
					// ────────────────────────────────────────────────────────────────
					// Extract Text from Binary File
					// ────────────────────────────────────────────────────────────────
					const binaryPropertyName = this.getNodeParameter('binaryPropertyName', i) as string;
					const outputFormat = this.getNodeParameter('outputFormat', i) as string;
					const additionalFields = this.getNodeParameter('additionalFields', i, {}) as any;
					const cleanupFields = this.getNodeParameter('cleanupOptions', i, {}) as any;

					const pdfPageRange = (additionalFields.pdfPageRange as string) || '';
					const excelSheetName = (additionalFields.excelSheetName as string) || '';
					const excelHeaderRow = additionalFields.excelHeaderRow !== false;
					const includeCharCounts = !!additionalFields.includeCharCounts;
					const maxFileSize = (additionalFields.maxFileSize as number) ?? 30;

					// Cleanup options — all default to true
					const cleanupOptions: CleanupOptions = {
						removeBase64:       cleanupFields.removeBase64       !== false,
						removeStyles:       cleanupFields.removeStyles       !== false,
						removeScripts:      cleanupFields.removeScripts      !== false,
						removeComments:     cleanupFields.removeComments     !== false,
						removeEmptyTags:    cleanupFields.removeEmptyTags    !== false,
						collapseWhitespace: cleanupFields.collapseWhitespace !== false,
					};

					const item = items[i];
					if (!item.binary || !item.binary[binaryPropertyName]) {
						throw new NodeOperationError(
							this.getNode(),
							`No binary property found with name '${binaryPropertyName}'.`,
							{ itemIndex: i }
						);
					}

					const binaryData = item.binary[binaryPropertyName];
					const buffer = await this.helpers.getBinaryDataBuffer(i, binaryPropertyName);
					const fileSizeBytes = buffer.length;
					const fileSizeMb = fileSizeBytes / (1024 * 1024);

					if (fileSizeMb > maxFileSize) {
						throw new NodeOperationError(
							this.getNode(),
							`The file size (${fileSizeMb.toFixed(2)} MB) exceeds the maximum allowed limit of ${maxFileSize} MB.`,
							{ itemIndex: i }
						);
					}

					// Detect file extension / type
					const ext = (
						binaryData.fileExtension ||
						binaryData.fileName?.split('.').pop() ||
						''
					).toLowerCase();

					let rawExtractedText = '';
					let formattedText = '';

					if (ext === 'pdf') {
						// ─── PDF PARSING ───
						const parseOptions: any = {};
						if (pdfPageRange) {
							parseOptions.pagerender = async function (pageData: any) {
								const pageNum = pageData.pageNumber;
								if (!isPageInRange(pageNum, pdfPageRange)) {
									return '';
								}
								const textContent = await pageData.getTextContent({ normalizeWhitespace: true });
								let lastY = '', text = '';
								for (const item of textContent.items) {
									if (lastY === item.transform[5] || !lastY) {
										text += item.str;
									} else {
										text += '\n' + item.str;
									}
									lastY = item.transform[5];
								}
								return text;
							};
						}

						const pdfData = await pdfParse(buffer, parseOptions);
						rawExtractedText = pdfData.text || '';

						if (outputFormat === 'html') {
							// Convert plain text lines to basic HTML paragraphs
							formattedText = rawExtractedText
								.split('\n')
								.map(line => line.trim() ? `<p>${line.trim()}</p>` : '')
								.filter(Boolean)
								.join('\n');
						} else if (outputFormat === 'markdown') {
							// Basic markdown formatting: double newlines for paragraphs
							formattedText = rawExtractedText
								.split('\n')
								.map(line => line.trim())
								.join('\n\n');
						} else {
							formattedText = rawExtractedText;
						}

					} else if (ext === 'docx') {
						// ─── DOCX PARSING ───
						if (outputFormat === 'html') {
							const result = await mammoth.convertToHtml({ buffer });
							formattedText = result.value;
							rawExtractedText = (await mammoth.extractRawText({ buffer })).value;
						} else if (outputFormat === 'markdown') {
							const result = await mammoth.convertToHtml({ buffer });
							formattedText = turndownService.turndown(result.value);
							rawExtractedText = (await mammoth.extractRawText({ buffer })).value;
						} else {
							const result = await mammoth.extractRawText({ buffer });
							formattedText = result.value;
							rawExtractedText = result.value;
						}

					} else if (ext === 'doc') {
						// ─── DOC PARSING (Legacy) ───
						const extractor = new WordExtractor();
						const doc = await extractor.extract(buffer);
						rawExtractedText = doc.getBody();

						if (outputFormat === 'html') {
							formattedText = rawExtractedText
								.split('\n')
								.map(line => line.trim() ? `<p>${line.trim()}</p>` : '')
								.filter(Boolean)
								.join('\n');
						} else if (outputFormat === 'markdown') {
							formattedText = rawExtractedText
								.split('\n')
								.map(line => line.trim())
								.join('\n\n');
						} else {
							formattedText = rawExtractedText;
						}

					} else if (['xlsx', 'xls', 'csv', 'ods'].includes(ext)) {
						// ─── SPREADSHEETS PARSING ───
						const workbook = XLSX.read(buffer, { type: 'buffer' });
						const sheetNames = excelSheetName
							? [excelSheetName]
							: workbook.SheetNames;

						// Check if specified sheet exists
						if (excelSheetName && !workbook.SheetNames.includes(excelSheetName)) {
							throw new NodeOperationError(
								this.getNode(),
								`The spreadsheet does not contain a sheet named '${excelSheetName}'.`,
								{ itemIndex: i }
							);
						}

						const partsRaw: string[] = [];
						const partsFormatted: string[] = [];

						for (const sheetName of sheetNames) {
							const sheet = workbook.Sheets[sheetName];
							if (!sheet) continue;

							// Raw text is always parsed as plain CSV
							const csv = XLSX.utils.sheet_to_csv(sheet);
							partsRaw.push(`--- Sheet: ${sheetName} ---\n${csv}`);

							if (outputFormat === 'markdown') {
								const mdTable = sheetToMarkdownTable(sheet, excelHeaderRow);
								partsFormatted.push(`### Sheet: ${sheetName}\n\n${mdTable}`);
							} else if (outputFormat === 'html') {
								const htmlTable = XLSX.utils.sheet_to_html(sheet);
								partsFormatted.push(`<h3>Sheet: ${sheetName}</h3>\n${htmlTable}`);
							} else {
								partsFormatted.push(`--- Sheet: ${sheetName} ---\n${csv}`);
							}
						}

						rawExtractedText = partsRaw.join('\n\n');
						formattedText = partsFormatted.join('\n\n');

					} else {
						// ─── FALLBACK TO PLAIN TEXT ───
						rawExtractedText = buffer.toString('utf-8');
						if (outputFormat === 'html') {
							formattedText = `<pre>${rawExtractedText}</pre>`;
						} else if (outputFormat === 'markdown') {
							formattedText = '```text\n' + rawExtractedText + '\n```';
						} else {
							formattedText = rawExtractedText;
						}
					}

					// Apply output cleanup
					formattedText = applyCleanup(formattedText, cleanupOptions);

					// Build execution result JSON
					const resultJson: any = {
						text: formattedText,
						fileType: ext,
						fileName: binaryData.fileName || '',
					};

					if (includeCharCounts) {
						resultJson.metrics = {
							originalCharCount: rawExtractedText.length,
							convertedCharCount: formattedText.length,
							fileSizeBytes,
						};
					}

					returnData.push({
						json: resultJson,
						pairedItem: { item: i },
					});

				} else if (operation === 'convertToFile') {
					// ────────────────────────────────────────────────────────────────
					// Convert JSON/Text to Binary File
					// ────────────────────────────────────────────────────────────────
					const fileFormat = this.getNodeParameter('fileFormat', i) as string;
					const sourceKey = this.getNodeParameter('sourceKey', i) as string;
					const fileNameInput = this.getNodeParameter('fileName', i) as string;
					const binaryPropertyNameOutput = this.getNodeParameter('binaryPropertyNameOutput', i) as string;
					const additionalFields = this.getNodeParameter('additionalFields', i, {}) as any;

					const excelSheetName = (additionalFields.excelSheetNameOutput as string) || 'Sheet1';
					const csvDelimiter = (additionalFields.csvDelimiter as string) || ',';

					const item = items[i];
					const sourceValue = item.json[sourceKey];

					if (sourceValue === undefined || sourceValue === null) {
						throw new NodeOperationError(
							this.getNode(),
							`No field named '${sourceKey}' found in the item's JSON data.`,
							{ itemIndex: i }
						);
					}

					let buffer: Buffer;
					let mimeType = 'application/octet-stream';
					let fileExt = fileFormat;

					if (fileFormat === 'xlsx' || fileFormat === 'csv') {
						// ─── GENERATE SPREADSHEETS ───
						let sheetData: any[];

						if (typeof sourceValue === 'string') {
							try {
								sheetData = JSON.parse(sourceValue);
							} catch (e) {
								// If it is just a string, parse lines by delimiter or treat it as a single row
								sheetData = [[sourceValue]];
							}
						} else if (Array.isArray(sourceValue)) {
							sheetData = sourceValue;
						} else {
							sheetData = [sourceValue];
						}

						const ws = Array.isArray(sheetData[0])
							? XLSX.utils.aoa_to_sheet(sheetData)
							: XLSX.utils.json_to_sheet(sheetData);

						const wb = XLSX.utils.book_new();
						XLSX.utils.book_append_sheet(wb, ws, excelSheetName);

						if (fileFormat === 'xlsx') {
							buffer = XLSX.write(wb, { bookType: 'xlsx', type: 'buffer' });
							mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
						} else {
							const csvContent = XLSX.utils.sheet_to_csv(ws, { FS: csvDelimiter });
							buffer = Buffer.from(csvContent, 'utf-8');
							mimeType = 'text/csv';
						}

					} else {
						// ─── GENERATE TEXT / HTML FILES ───
						let textContent = '';
						if (typeof sourceValue === 'object') {
							textContent = JSON.stringify(sourceValue, null, 2);
						} else {
							textContent = String(sourceValue);
						}

						buffer = Buffer.from(textContent, 'utf-8');

						if (fileFormat === 'html') {
							mimeType = 'text/html';
						} else {
							mimeType = 'text/plain';
							fileExt = 'txt';
						}
					}

					const finalFileName = `${fileNameInput}.${fileExt}`;
					const binaryData = await this.helpers.prepareBinaryData(buffer, finalFileName, mimeType);

					// We do a deep copy of the item and append binary field, preserving standard performance rules
					const newItem = JSON.parse(JSON.stringify(item));
					newItem.binary = newItem.binary || {};
					newItem.binary[binaryPropertyNameOutput] = binaryData;

					returnData.push(newItem);
				}

			} catch (error) {
				if (this.continueOnFail()) {
					const errorMessage = error instanceof Error ? error.message : String(error);
					returnData.push({
						json: { error: errorMessage },
						pairedItem: { item: i },
					});
					continue;
				}
				throw error;
			}
		}

		return [returnData];
	}
}
