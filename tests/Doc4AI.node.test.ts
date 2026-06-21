import { isPageInRange, getColName, sheetToMarkdownTable } from '../nodes/Doc4AI/Doc4AI.node';
import * as XLSX from 'xlsx';

describe('Doc4AI Helper Functions', () => {
	test('isPageInRange', () => {
		expect(isPageInRange(1, '1-3')).toBe(true);
		expect(isPageInRange(2, '1-3')).toBe(true);
		expect(isPageInRange(4, '1-3')).toBe(false);
		expect(isPageInRange(5, '1-3, 5, 8-10')).toBe(true);
		expect(isPageInRange(6, '1-3, 5, 8-10')).toBe(false);
		expect(isPageInRange(8, '1-3, 5, 8-10')).toBe(true);
		expect(isPageInRange(9, '1-3, 5, 8-10')).toBe(true);
		expect(isPageInRange(11, '')).toBe(true);
		expect(isPageInRange(11, '   ')).toBe(true);
	});

	test('getColName', () => {
		expect(getColName(0)).toBe('A');
		expect(getColName(1)).toBe('B');
		expect(getColName(25)).toBe('Z');
		expect(getColName(26)).toBe('AA');
		expect(getColName(27)).toBe('AB');
	});

	test('sheetToMarkdownTable with headers', () => {
		const ws = XLSX.utils.aoa_to_sheet([
			['Name', 'Age'],
			['Alice', 25],
			['Bob', 30]
		]);
		const md = sheetToMarkdownTable(ws, true);
		expect(md).toContain('| Name | Age |');
		expect(md).toContain('| Alice | 25 |');
		expect(md).toContain('| Bob | 30 |');
	});

	test('sheetToMarkdownTable without headers', () => {
		const ws = XLSX.utils.aoa_to_sheet([
			['Alice', 25],
			['Bob', 30]
		]);
		const md = sheetToMarkdownTable(ws, false);
		expect(md).toContain('| A | B |');
		expect(md).toContain('| Alice | 25 |');
		expect(md).toContain('| Bob | 30 |');
	});
});

describe('Doc4AI Execute - File Size Limit', () => {
	const { Doc4AI } = require('../nodes/Doc4AI/Doc4AI.node');
	let doc4aiNode: any;

	beforeEach(() => {
		doc4aiNode = new Doc4AI();
	});

	test('should reject file if size exceeds maxFileSize limit (default 30MB)', async () => {
		const mockContext: any = {
			getInputData: () => [{ json: {}, binary: { data: { fileName: 'large.txt' } } }],
			getNodeParameter: (paramName: string, itemIndex: number, fallback?: any) => {
				if (paramName === 'operation') return 'extractText';
				if (paramName === 'binaryPropertyName') return 'data';
				if (paramName === 'outputFormat') return 'text';
				if (paramName === 'additionalFields') return fallback ?? {};
				if (paramName === 'cleanupOptions') return fallback ?? {};
				return undefined;
			},
			helpers: {
				getBinaryDataBuffer: async () => Buffer.alloc(31 * 1024 * 1024), // 31 MB
			},
			getNode: () => ({ name: 'Doc4AI' }),
			continueOnFail: () => false,
		};

		await expect(doc4aiNode.execute.call(mockContext)).rejects.toThrow(
			'The file size (31.00 MB) exceeds the maximum allowed limit of 30 MB.'
		);
	});

	test('should accept file if size is within maxFileSize limit (custom 50MB)', async () => {
		const mockContext: any = {
			getInputData: () => [{ json: {}, binary: { data: { fileName: 'large.txt' } } }],
			getNodeParameter: (paramName: string, itemIndex: number, fallback?: any) => {
				if (paramName === 'operation') return 'extractText';
				if (paramName === 'binaryPropertyName') return 'data';
				if (paramName === 'outputFormat') return 'text';
				if (paramName === 'additionalFields') return { maxFileSize: 50 };
				if (paramName === 'cleanupOptions') return fallback ?? {};
				return undefined;
			},
			helpers: {
				getBinaryDataBuffer: async () => Buffer.alloc(31 * 1024 * 1024), // 31 MB
			},
			getNode: () => ({ name: 'Doc4AI' }),
			continueOnFail: () => false,
		};

		const result = await doc4aiNode.execute.call(mockContext);
		expect(result).toBeDefined();
		expect(result[0][0].json.text).toBeDefined();
	});
});

