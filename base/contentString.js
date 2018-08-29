define([], function () {
    var ContentString = function (template) {
        if (template) this.template = template;
    };

    ContentString.prototype = function () {
        var format = function(){
            var args = arguments;
            this.template = this.template.replace(/{(\d+)}/g, function(match, number){
                return typeof args[number] !== "undefined" ? args[number] : match;
            });
            return this.template;
        };
        
        var toString = function(){
            return this.template;
        }

        return { format: format, toString: toString };
    }();

    return ContentString;
});