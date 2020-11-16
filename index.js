const gm = require('gm');
const xlsx = require('node-xlsx');
const fs = require('fs');
const compressing = require('compressing');

const PINGFANG = './src/font/PingFang_medium.ttf';
const HEITI = './src/font/STHEITI.ttf';

const NormalTemplate = './src/template/normal.jpg';
const SHEJIZHIXINGTemplate = './src/template/设计之星.jpg';
const ZISHENTemplate = './src/template/资深达人.jpg';
const SHOUXITemplate = './src/template/首席达人.jpg';

let sum = 0;
let croppedNum = 0;

const clear = (path) => {
    let files = [];
    if (fs.existsSync(path)) {
        files = fs.readdirSync(path);
        files.forEach((file, index) => {
            let curPath = path + "/" + file;
            if (fs.statSync(curPath).isDirectory()) {
                clear(curPath); //递归删除文件夹
            } else {
                fs.unlinkSync(curPath); //删除文件
            }
        });
        fs.rmdirSync(path);
    }
}

const compress = () => {
    compressing.zip.compressDir('./result', './result.zip')
        .then(() => {
            console.log('Compression success!');
        })
        .catch(err => {
            console.error(err);
        });
}

const croppedCallback = (level, name) => {
    croppedNum++;
    console.log(`${level}:${name} cropped`);
    if (croppedNum === sum) {
        console.log('Cropped over, start compress...')
        compress();
    }
}

const drawByNormal = (pathBase, template, name, level, levelName, index) => {
    gm(template)
        .font(PINGFANG)
        .fontSize(145)
        .fill('#030000')
        .drawText(0, -420, name, 'Center')
        .fontSize(140)
        .fill('#b62d2b')
        .drawText(0 , 480, `${levelName}设计师` , 'Center')
        .quality(100)
        .write(`${pathBase + level}/${index+1}-${name}.jpg`, err => {
            if (!err) {
                croppedCallback(level, name);
            } else {
                console.log(err)
            }
        });
}

const drawBySpecial = (pathBase, template, name, level, index, color) => {
    gm(template)
        .font(HEITI)
        .fontSize(145)
        .fill(color)
        .drawText(330, -180, name, 'Center')
        .quality(100)
        .write(`${pathBase + level}/${index+1}-${name}.jpg`, err => {
            if (!err) {
                croppedCallback(level, name);
            } else {
                console.log(err)
            }
        });
}

const crop = (xlsxPath, templateJpg) => {
    clear('./result')
    const data = xlsx.parse(xlsxPath)
    const pathBase = './result/'
    if (!fs.existsSync(pathBase)) {
        fs.mkdirSync(pathBase);
    }
    data.forEach(({
        name: level,
        data: list
    }) => {
        list.shift();
        const path = pathBase + level;
        if (!fs.existsSync(path)) {
            fs.mkdirSync(path);
        }
        list.forEach(([name], index) => {
            if (level && name) {
                sum++;
                switch (level) {
                    case '资深':
                        drawBySpecial(pathBase, ZISHENTemplate, name, level, index, '#313536');
                        break;
                    case '首席':
                        drawBySpecial(pathBase, SHOUXITemplate, name, level, index, '#313536');
                        break;
                    case '设计之星':
                        drawByNormal(pathBase, SHEJIZHIXINGTemplate, name, level, '设计之星', index);
                        break;
                    default:
                        drawByNormal(pathBase, NormalTemplate, name, level, level, index);
                        return;
                }
            }
        });
    });
}

crop('./src/list/证书制作名单20201108.xlsx')