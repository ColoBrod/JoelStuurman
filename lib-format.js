function formatDate(date, short = false) {
  let d = date.getDate();
  let m = date.getMonth();
  let y = date.getFullYear();
  switch (m) {
    case 0:
      m = "January";
      break;
    case 1:
      m = "February";
      break;
    case 2:
      m = "March";
      break;
    case 3:
      m = "April";
      break;
    case 4:
      m = "May";
      break;
    case 5:
      m = "June";
      break;
    case 6:
      m = "July";
      break;
    case 7:
      m = "August";
      break;
    case 8:
      m = "September";
      break;
    case 9:
      m = "October";
      break;
    case 10:
      m = "November";
      break;
    case 11:
      m = "December";
      break;
  }
  if (short) m = m.substring(0,3);
  return `${m} ${d}, ${y}`;
}

function humanize(s) {

  // System for American Numbering 
  var th_val = ['', 'thousand', 'million', 'billion', 'trillion'];
  // System for uncomment this line for Number of English 
  // var th_val = ['','thousand','million', 'milliard','billion'];
  
  var dg_val = ['zero', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine'];
  var tn_val = ['ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen', 'seventeen', 'eighteen', 'nineteen'];
  var tw_val = ['twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty', 'ninety'];

  s = s.toString();
    s = s.replace(/[\, ]/g, '');
    if (s != parseFloat(s))
        return 'not a number ';
    var x_val = s.indexOf('.');
    if (x_val == -1)
        x_val = s.length;
    if (x_val > 15)
        return 'too big';
    var n_val = s.split('');
    var str_val = '';
    var sk_val = 0;
    for (var i = 0; i < x_val; i++) {
        if ((x_val - i) % 3 == 2) {
            if (n_val[i] == '1') {
                str_val += tn_val[Number(n_val[i + 1])] + ' ';
                i++;
                sk_val = 1;
            } else if (n_val[i] != 0) {
                str_val += tw_val[n_val[i] - 2] + ' ';
                sk_val = 1;
            }
        } else if (n_val[i] != 0) {
            str_val += dg_val[n_val[i]] + ' ';
            if ((x_val - i) % 3 == 0)
                str_val += 'hundred ';
            sk_val = 1;
        }
        if ((x_val - i) % 3 == 1) {
            if (sk_val)
                str_val += th_val[(x_val - i - 1) / 3] + ' ';
            sk_val = 0;
        }
    }
    if (x_val != s.length) {
        var y_val = s.length;
        str_val += 'point ';
        for (var i = x_val + 1; i < y_val; i++)
            str_val += dg_val[n_val[i]] + ' ';
    }
    return str_val.replace(/\s+/g, ' ');
}

// function humanize(num) {
//   var ones = ['', 'one', 'two', 'three', 'four', 'five', 'six', 'seven', 'eight', 'nine',
//               'ten', 'eleven', 'twelve', 'thirteen', 'fourteen', 'fifteen', 'sixteen',
//               'seventeen', 'eighteen', 'nineteen'];
//   var tens = ['', '', 'twenty', 'thirty', 'forty', 'fifty', 'sixty', 'seventy', 'eighty',
//               'ninety'];

//   var numString = num.toString();

//   if (num < 0) throw new Error('Negative numbers are not supported.');

//   if (num === 0) return 'zero';

//   //the case of 1 - 20
//   if (num < 20) {
//     return ones[num];
//   }

//   if (numString.length === 2) {
//     return tens[numString[0]] + ' ' + ones[numString[1]];
//   }

//   //100 and more
//   if (numString.length == 3) {
//     if (numString[1] === '0' && numString[2] === '0')
//       return ones[numString[0]] + ' hundred';
//     else
//       return ones[numString[0]] + ' hundred and ' + humanize(+(numString[1] + numString[2]));
//   }

//   if (numString.length === 4) {
//     var end = +(numString[1] + numString[2] + numString[3]);
//     if (end === 0) return ones[numString[0]] + ' thousand';
//     if (end < 100) return ones[numString[0]] + ' thousand and ' + humanize(end);
//     return ones[numString[0]] + ' thousand ' + humanize(end);
//   }
// }

function numberWithCommas(x) {
  return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ",") + ".00";
}