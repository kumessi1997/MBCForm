function doExportPdf(filename) {
    var element = document.querySelector('.page-wrap');

    // Tạm bỏ padding top/bottom và box-shadow trước khi render
    // để html2pdf kiểm soát lề đồng đều trên tất cả các trang
    var prevPadding   = element.style.padding;
    var prevBoxShadow = element.style.boxShadow;
    element.style.padding   = '0 40px';
    element.style.boxShadow = 'none';

    // Đổi header cells sang nền trắng / chữ đen để tối ưu khi in
    var darkCells = element.querySelectorAll('.cell-num, .row-section td');
    var darkCellsPrev = Array.prototype.map.call(darkCells, function (cell) {
        return {
            background:      cell.style.background,
            backgroundColor: cell.style.backgroundColor,
            color:           cell.style.color,
            fontWeight:      cell.style.fontWeight,
            borderColor:     cell.style.borderColor
        };
    });
    darkCells.forEach(function (cell) {
        cell.style.background      = '#fff';
        cell.style.backgroundColor = '#fff';
        cell.style.color           = '#000';
        cell.style.fontWeight      = 'bold';
    });

    // 26px (padding gốc) / 96dpi * 25.4 ≈ 6.9mm → dùng 7mm cho tất cả các trang
    var opt = {
        margin:      [7, 0, 7, 0],
        filename:    filename,
        image:       { type: 'jpeg', quality: 0.98 },
        html2canvas: {
            scale:           3,
            useCORS:         true,
            logging:         false,
            backgroundColor: '#ffffff',
            scrollX:         0,
            scrollY:         0
        },
        jsPDF: {
            unit:        'mm',
            format:      'a4',
            orientation: 'portrait'
        },
        pagebreak: {
            mode:   ['css', 'legacy'],
            before: '.pdf-page-break',
            avoid:  'tr'
        }
    };

    html2pdf().set(opt).from(element).save().then(function () {
        // Khôi phục style sau khi xuất xong
        element.style.padding   = prevPadding;
        element.style.boxShadow = prevBoxShadow;

        darkCells.forEach(function (cell, i) {
            cell.style.background      = darkCellsPrev[i].background;
            cell.style.backgroundColor = darkCellsPrev[i].backgroundColor;
            cell.style.color           = darkCellsPrev[i].color;
            cell.style.fontWeight      = darkCellsPrev[i].fontWeight;
        });
    });
}

function doExportDocx(filename) {
    var wrap = document.querySelector('.page-wrap');
    var clone = wrap.cloneNode(true);

    var liveInputs = wrap.querySelectorAll('input');
    var cloneInputs = clone.querySelectorAll('input');
    liveInputs.forEach(function(input, i) {
        var ci = cloneInputs[i];
        if (!ci) return;
        if (input.type === 'radio' || input.type === 'checkbox') {
            var mark = document.createElement('span');
            mark.textContent = input.checked
                ? (input.type === 'checkbox' ? '☑' : '●')
                : (input.type === 'checkbox' ? '☐' : '○');
            ci.parentNode.replaceChild(mark, ci);
        } else {
            var span = document.createElement('span');
            span.style.cssText = 'border-bottom:1px solid #666;display:inline-block;width:100%;min-width:60px;';
            span.textContent = input.value || ' ';
            ci.parentNode.replaceChild(span, ci);
        }
    });

    clone.querySelectorAll('svg').forEach(function(s) { s.remove(); });

    var css = [
        'body{font-family:Arial,sans-serif;font-size:9pt;color:#000;margin:1cm 1.5cm;}',
        'table{width:100%;border-collapse:collapse;font-size:8.5pt;}',
        'td,th{border:1px solid #000;padding:3px 5px;vertical-align:top;line-height:1.35;}',
        '.cell-num,.row-section td{background:#1a3a6b;color:#fff;font-weight:bold;}',
        '.cell-num{text-align:center;width:18px;}',
        '.row-section-center td{text-align:center;}',
        '.lbl{font-weight:bold;font-size:8pt;display:block;}',
        '.lbl-en{font-weight:normal;font-style:italic;font-size:7.5pt;}',
        '.lbl-only{font-size:8pt;display:block;}',
        '.lbl-only-en{font-style:italic;font-size:7.5pt;}',
        '.sig-h90{height:88px;vertical-align:top;text-align:center;padding:4px 6px;}',
        '.sig-h70{height:72px;vertical-align:top;text-align:center;padding:4px 6px;}',
        '.sig-title{font-weight:bold;font-size:8pt;}',
        '.sig-sub{font-style:italic;font-size:7.5pt;}',
        '.doc-footer{font-size:7pt;color:#333;margin-top:6px;border-top:1px solid #999;padding-top:4px;}',
        '.commit-block{font-size:7.5pt;line-height:1.5;padding:4px 6px;border:1px solid #000;margin-top:-1px;}',
        '.doc-title h1{font-size:11pt;font-weight:bold;color:#1a3a6b;text-transform:uppercase;}',
        '.doc-title h2{font-size:10pt;color:#1a3a6b;font-weight:bold;font-style:italic;}',
        '.doc-title h3{font-size:8.5pt;color:#444;font-style:italic;}'
    ].join('');

    // MSWord XML HTML format — không cần thư viện ngoài, Word/LibreOffice đọc được
    var docHtml = '<html xmlns:o="urn:schemas-microsoft-com:office:office" '
        + 'xmlns:w="urn:schemas-microsoft-com:office:word" '
        + 'xmlns="http://www.w3.org/TR/REC-html40">'
        + '<head><meta charset="UTF-8">'
        + '<!--[if gte mso 9]><xml><w:WordDocument><w:View>Normal</w:View>'
        + '<w:Zoom>100</w:Zoom><w:DoNotOptimizeForBrowser/>'
        + '<\/w:WordDocument><\/xml><![endif]-->'
        + '<style>' + css + '<\/style>'
        + '<\/head><body>'
        + clone.innerHTML
        + '<\/body><\/html>';

    var blob = new Blob(['﻿', docHtml], { type: 'application/msword' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}
