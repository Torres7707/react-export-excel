import * as React from 'react';
import { Button, Table, Upload } from 'antd';
import * as XLSX from 'xlsx';
import { WritingOptions } from 'xlsx';

export interface IExcelState {
	columns?: Object;
	data?: Object;
	fileList?: [];
}

class Excel extends React.Component<IExcelState> {
	readonly state = {
		columns: [
			{ title: '姓名', dataIndex: 'name' },
			{ title: '年龄', dataIndex: 'age' },
			{ title: '地址', dataIndex: 'address' },
		],
		data: [],
		fileList: [],
	};

	formatTitleOrFileld = (a, b) => {
		const entozh = {};
		this.state.columns.forEach((item) => {
			entozh[item[a]] = item[b];
		});
		return entozh;
	};

	handleImpotedJson = (array, file) => {
		const header = array[0]; // 表格头部，title
		console.log('header', header); // ["姓名", "年龄", "地址"]
		const entozh = this.formatTitleOrFileld('title', 'dataIndex');
		// { '姓名':'name','年龄':'age','地址':'address'}
		console.log('entozh', entozh);
		const firstRow = header.map((item) => entozh[item]);
		console.log('firstRow', firstRow); // ["name", "age", "address"]

		const newArray = [...array];
		console.log('newArray', newArray);

		newArray.splice(0, 1); // [["呜呜呜呜", 111, "啊啊啊啊啊"]]

		const json: any = newArray.map((item, index) => {
			const newitem = {};
			item.forEach((im, i) => {
				const newKey = firstRow[i] || i;
				newitem[newKey] = im;
			});
			return newitem;
		});
		console.log('json', json); // [{name: "呜呜呜呜", age: 111, address: "啊啊啊啊啊"}]
		const formatData = json.map((item) => ({
			name: item.name,
			age: item.age,
			address: item.address,
		}));

		console.log('formatData', formatData); // [{name: "呜呜呜呜", age: 111, address: "啊啊啊啊啊"}]

		this.setState({ data: formatData, fileList: [file] });

		return formatData;
	};

	sheet2blob = (sheet, sheetName) => {
		sheetName = sheetName || 'sheet1';
		var workbook = {
			SheetNames: [sheetName],
			Sheets: {},
		};
		workbook.Sheets[sheetName] = sheet; // 生成excel的配置项

		var wopts: WritingOptions = {
			type: 'binary',
			bookSST: false, // 是否生成Shared String Table，官方解释是，如果开启生成速度会下降，但在低版本IOS设备上有更好的兼容性
			bookType: 'xlsx', // 要生成的文件类型
		};
		var wbout = XLSX.write(workbook, wopts);
		var blob = new Blob([s2ab(wbout)], {
			type: 'application/octet-stream',
		}); // 字符串转ArrayBuffer
		function s2ab(s) {
			var buf = new ArrayBuffer(s.length);
			var view = new Uint8Array(buf);
			for (var i = 0; i != s.length; ++i) view[i] = s.charCodeAt(i) & 0xff;
			return buf;
		}
		return blob;
	};

	openDownloadDialog = (url, saveName) => {
		if (typeof url == 'object' && url instanceof Blob) {
			url = URL.createObjectURL(url); // 创建blob地址
		}
		var aLink = document.createElement('a');
		aLink.href = url;
		aLink.download = saveName || ''; // HTML5新增的属性，指定保存文件名，可以不要后缀，注意，file:///模式下不会生效
		var event;
		if (window.MouseEvent) event = new MouseEvent('click');
		else {
			event = document.createEvent('MouseEvents');
			event.initMouseEvent(
				'click',
				true,
				false,
				window,
				0,
				0,
				0,
				0,
				0,
				false,
				false,
				false,
				false,
				0,
				null
			);
		}
		aLink.dispatchEvent(event);
	};

	handleExportAll = (e) => {
		const entozh = {
			name: '姓名',
			age: '年龄',
			address: '地址',
		};

		const nowdata = this.state.data;

		const json = nowdata.map((item) => {
			return Object.keys(item).reduce((newData, key) => {
				const newKey = entozh[key] || key;
				newData[newKey] = item[key];
				return newData;
			}, {});
		});

		console.log('json', json); // [{姓名: "呜呜呜呜", 年龄: 111, 地址: "啊啊啊啊啊"}]

		const sheet = XLSX.utils.json_to_sheet(json);

		this.openDownloadDialog(this.sheet2blob(sheet, undefined), `全部信息.xlsx`);
	};

	handleExportDocument = (e) => {
		const entozh = {
			name: '姓名',
			age: '年龄',
			address: '地址',
		};

		let nowdata = [{ name: '' }, { age: '' }, { address: '' }];

		const json = nowdata.map((item) => {
			console.log('11', Object.keys(item));
			return Object.keys(item).reduce((newData, key) => {
				const newKey = entozh[key] || key;
				newData[newKey] = item[key];
				return newData;
			}, {});
		});
		console.log('11', json);

		const sheet = XLSX.utils.json_to_sheet(json);

		// this.openDownloadDialog(
		// 	this.sheet2blob(sheet, undefined),
		// 	`标准格式文件.xlsx`
		// );
	};

	render() {
		const { columns, data, fileList } = this.state;

		const uploadProps = {
			onRemove: (file) => {
				this.setState((state) => ({
					data: [],
					fileList: [],
				}));
			},
			accept: '.xls,.xlsx,application/vnd.ms-excel',
			beforeUpload: (file) => {
				const _this = this;
				const f = file;
				const reader = new FileReader();
				reader.onload = function (e) {
					const datas = e.target.result;
					const workbook = XLSX.read(datas, {
						type: 'binary',
					}); //尝试解析datas
					console.log('workbook', workbook);
					console.log('workbook.sheetName', workbook.SheetNames);
					console.log('workbook.sheets', workbook.Sheets);

					const first_worksheet = workbook.Sheets[workbook.SheetNames[0]]; //是工作簿中的工作表的有序列表

					const jsonArr = XLSX.utils.sheet_to_json(first_worksheet, {
						header: 1,
					}); //将工作簿对象转换为JSON对象数组
					console.log('jsonArr', jsonArr);
					_this.handleImpotedJson(jsonArr, file);
				};
				reader.readAsBinaryString(f);
				return false;
			},
			fileList,
		};

		return (
			<div>
				<Upload {...uploadProps}>
					<Button type="primary">Excel导入</Button>
				</Upload>

				<Button type="primary" onClick={this.handleExportAll}>
					Excel导出数据
				</Button>

				<Button type="primary" onClick={this.handleExportDocument}>
					Excel导出格式文件
				</Button>

				<Table columns={columns} dataSource={data} bordered></Table>
			</div>
		);
	}
}

export default Excel;
