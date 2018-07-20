define([], function () {
    sortObject = function (obj) {
        if (typeof obj !== 'object' || obj.length)
            return obj;
        var temp = {};
        var keys = [];
        for (var key in obj)
            keys.push(key);
        keys.sort();
        for (var index in keys)
            temp[keys[index]] = sortObject(obj[keys[index]]);
        return temp;
    };

    var stringify = function (obj) {
        return JSON.stringify(sortObject(obj));
    };

    return {
        stringify: stringify
    };
});