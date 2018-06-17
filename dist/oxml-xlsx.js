"use strict";

(function(window){
    createXLSX = function(){
        var _xlsx = {};

    };
    
    if (!window.oxml) {
        window.oxml = {};
    }

    window.oxml.createXLSX = createXLSX;
})(window);