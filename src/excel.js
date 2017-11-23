import sax from 'sax';
import unzip from 'unzip2';
import fs from 'fs-extra';
import _ from 'lodash';
export default filename => new Excel(filename);
class Excel 
{
	constructor(filename)
	{
		this.filename = filename;
	}
	book = () => new Promise(resolve =>
	{
		fs.createReadStream(this.filename).pipe(unzip.Parse())
		.on('entry', async entry =>
		{
			switch (entry.path)
			{
				case 'xl/workbook.xml': 
					this.sheets = await parseSheet(entry);
					break;
				case 'xl/_rels/workbook.xml.rels': 
					this.rels = await parseRels(entry);
					break;
				case 'xl/sharedStrings.xml':
					this.sst = await parseSst(entry);
					break;
			}
			entry.autodrain();
		})
		.on('finish', () =>
		{
			_(this.sheets).each(sheet =>
			{
				sheet.entry = this.rels[sheet.rId];
			});
			resolve(this);
		});
	});
}
const parseSheet = entry => new Promise(resolve =>
{
	const sheets = [];
	parseSax(entry)
	.on('node', node =>
	{
		if (node.path === 'workbook/sheets/sheet')
		{
			const { name, sheetId, 'r:id': rId } = node.attributes;
			sheets.push({ name, sheetId, rId });
		}
	})
	.on('end', () =>
	{
		resolve(sheets);
	});
});
const parseRels = entry => new Promise(resolve =>
{
	const rels = {};
	parseSax(entry)
	.on('node', node =>
	{
		if (node.path === 'Relationships/Relationship')
		{
			const { Id: id, Target: entry } = node.attributes;
			rels[id] = entry;
		}
	})
	.on('end', () =>
	{
		resolve(rels);
	});
});
const parseSst = entry => new Promise(resolve =>
{
	const sst = [];
	parseSax(entry)
	.on('node', node =>
	{
		if (node.path === 'sst/si/t')
		{
			sst.push(node.text);
		}
	})
	.on('end', () =>
	{
		resolve(sst);
	});
});
const parseSax = entry =>
{
	let c = 10;
	const strict = true;  
	const trim = false; 
	const position = true; 
	const strictEntities = true; 
	const nodes = [];
	const names = [];
	const stream = entry.pipe(sax.createStream(strict, { trim, position, strictEntities }))
	stream.on('opentag', node => 
	{
		nodes.push(node);
		names.push(node.name);
	})
	.on('text', text =>
	{
		if (nodes.length === 0) return;
		_(nodes).last().text = text;
	})
	.on('closetag', name =>
	{
		const node = _(nodes).last();
		const path = _(names).join('/');
		stream.emit('node', { ...node, path });
		nodes.pop();
		names.pop();
	});
	return stream;
}; 