/**
 * Created by kin on 2017/2/8.
 * josn导出excel
 * mail：cuikangjie_90h@126.com
 */

(function(){
  'user static';
  var eje=function (option) {
      this.data=option.data || {};
      this.filter=option.filter;
      this.fileName=option.fileName || 'download';
      this.instance=function (exedata) {
          var data = exedata;
          var ws_name = "sheet";

          var wb = new Workbook(), ws = sheet_from_array_of_arrays(data);
          ws['!merges'] = [];
          wb.SheetNames.push(ws_name);
          wb.Sheets[ws_name] = ws;
          var wbout = XLSX.write(wb, {bookType:'xlsx', bookSST:false, type: 'binary'});
          this.saveAs(new Blob([s2ab(wbout)],{type:"application/octet-stream"}), this.fileName+".xlsx")
      };
      this.changeData=function (data,filter) {
          var re=[];
          typeof data[0][0]=='undefined' ? (function () {
             //对象
              filter ? (function () {
                  //存在过滤
                  data.forEach(function (obj) {
                      var one=[];
                      filter.forEach(function (no) {
                          one.push(obj[no]);
                      });
                      re.push(one);
                  });
              })() :(function () {
                  //不存在过滤
                  data.forEach(function (obj) {
                      var col=Object.keys(obj);
                      var one=[];
                      col.forEach(function (no) {
                          one.push(obj[no]);
                      });
                      re.push(one);
                  });

              })();
          })() : (function(){
             re= data;
          })();
          return re;
      };
      var Workbook= function() {
          if(!(this instanceof Workbook)) return new Workbook();
          this.SheetNames = [];
          this.Sheets = {};
      };
      var s2ab= function(s) {
          var buf = new ArrayBuffer(s.length);
          var view = new Uint8Array(buf);
          for (var i=0; i!=s.length; ++i) view[i] = s.charCodeAt(i) & 0xFF;
          return buf;
      };
      var datenum= function(v, date1904) {
          if(date1904) v+=1462;
          var epoch = Date.parse(v);
          return (epoch - new Date(Date.UTC(1899, 11, 30))) / (24 * 60 * 60 * 1000);
      };
      var sheet_from_array_of_arrays=function (data) {
          var ws = {};
          var range = {s: {c:10000000, r:10000000}, e: {c:0, r:0 }};
          for(var R = 0; R != data.length; ++R) {
              for(var C = 0; C != data[R].length; ++C) {
                  if(range.s.r > R) range.s.r = R;
                  if(range.s.c > C) range.s.c = C;
                  if(range.e.r < R) range.e.r = R;
                  if(range.e.c < C) range.e.c = C;
                  var cell = {v: data[R][C] };
                  if(cell.v == null) continue;
                  var cell_ref = XLSX.utils.encode_cell({c:C,r:R});

                  if(typeof cell.v === 'number') cell.t = 'n';
                  else if(typeof cell.v === 'boolean') cell.t = 'b';
                  else if(cell.v instanceof Date) {
                      cell.t = 'n'; cell.z = XLSX.SSF._table[14];
                      cell.v = datenum(cell.v);
                  }
                  else cell.t = 's';
                  ws[cell_ref] = cell;
              }
          }
          if(range.s.c < 10000000) ws['!ref'] = XLSX.utils.encode_range(range);
          return ws;
      };
      this.saveAs = this.saveAs
          // IE 10+ (native saveAs)
          || (typeof navigator !== "undefined" &&
          navigator.msSaveOrOpenBlob && navigator.msSaveOrOpenBlob.bind(navigator))
          // Everyone else
          || (function(view) {
              "use strict";
              // IE <10 is explicitly unsupported
              if (typeof navigator !== "undefined" &&
                  /MSIE [1-9]\./.test(navigator.userAgent)) {
                  return;
              }
              var
                  doc = view.document
                  // only get URL when necessary in case BlobBuilder.js hasn't overridden it yet
                  , get_URL = function() {
                      return view.URL || view.webkitURL || view;
                  }
                  , URL = view.URL || view.webkitURL || view
                  , save_link = doc.createElementNS("http://www.w3.org/1999/xhtml", "a")
                  , can_use_save_link = !view.externalHost && "download" in save_link
                  , click = function(node) {
                      var event = doc.createEvent("MouseEvents");
                      event.initMouseEvent(
                          "click", true, false, view, 0, 0, 0, 0, 0
                          , false, false, false, false, 0, null
                      );
                      node.dispatchEvent(event);
                  }
                  , webkit_req_fs = view.webkitRequestFileSystem
                  , req_fs = view.requestFileSystem || webkit_req_fs || view.mozRequestFileSystem
                  , throw_outside = function(ex) {
                      (view.setImmediate || view.setTimeout)(function() {
                          throw ex;
                      }, 0);
                  }
                  , force_saveable_type = "application/octet-stream"
                  , fs_min_size = 0
                  , deletion_queue = []
                  , process_deletion_queue = function() {
                      var i = deletion_queue.length;
                      while (i--) {
                          var file = deletion_queue[i];
                          if (typeof file === "string") { // file is an object URL
                              URL.revokeObjectURL(file);
                          } else { // file is a File
                              file.remove();
                          }
                      }
                      deletion_queue.length = 0; // clear queue
                  }
                  , dispatch = function(filesaver, event_types, event) {
                      event_types = [].concat(event_types);
                      var i = event_types.length;
                      while (i--) {
                          var listener = filesaver["on" + event_types[i]];
                          if (typeof listener === "function") {
                              try {
                                  listener.call(filesaver, event || filesaver);
                              } catch (ex) {
                                  throw_outside(ex);
                              }
                          }
                      }
                  }
                  , FileSaver = function(blob, name) {
                      // First try a.download, then web filesystem, then object URLs
                      var
                          filesaver = this
                          , type = blob.type
                          , blob_changed = false
                          , object_url
                          , target_view
                          , get_object_url = function() {
                              var object_url = get_URL().createObjectURL(blob);
                              deletion_queue.push(object_url);
                              return object_url;
                          }
                          , dispatch_all = function() {
                              dispatch(filesaver, "writestart progress write writeend".split(" "));
                          }
                          // on any filesys errors revert to saving with object URLs
                          , fs_error = function() {
                              // don't create more object URLs than needed
                              if (blob_changed || !object_url) {
                                  object_url = get_object_url(blob);
                              }
                              if (target_view) {
                                  target_view.location.href = object_url;
                              } else {
                                  if(navigator.userAgent.match(/7\.[\d\s\.]+Safari/)	// is Safari 7.x
                                      && typeof window.FileReader !== "undefined"			// can convert to base64
                                      && blob.size <= 1024*1024*150										// file size max 150MB
                                  ) {
                                      var reader = new window.FileReader();
                                      reader.readAsDataURL(blob);
                                      reader.onloadend = function() {
                                          var frame = doc.createElement("iframe");
                                          frame.src = reader.result;
                                          frame.style.display = "none";
                                          doc.body.appendChild(frame);
                                          dispatch_all();
                                          return;
                                      };
                                      filesaver.readyState = filesaver.DONE;
                                      filesaver.savedAs = filesaver.SAVEDASUNKNOWN;
                                      return;
                                  }
                                  else {
                                      window.open(object_url, "_blank");
                                      filesaver.readyState = filesaver.DONE;
                                      filesaver.savedAs = filesaver.SAVEDASBLOB;
                                      dispatch_all();
                                      return;
                                  }
                              }
                          }
                          , abortable = function(func) {
                              return function() {
                                  if (filesaver.readyState !== filesaver.DONE) {
                                      return func.apply(this, arguments);
                                  }
                              };
                          }
                          , create_if_not_found = {create: true, exclusive: false}
                          , slice
                          ;
                      filesaver.readyState = filesaver.INIT;
                      if (!name) {
                          name = "download";
                      }
                      if (can_use_save_link) {
                          object_url = get_object_url(blob);
                          // FF for Android has a nasty garbage collection mechanism
                          // that turns all objects that are not pure javascript into 'deadObject'
                          // this means `doc` and `save_link` are unusable and need to be recreated
                          // `view` is usable though:
                          doc = view.document;
                          save_link = doc.createElementNS("http://www.w3.org/1999/xhtml", "a");
                          save_link.href = object_url;
                          save_link.download = name;
                          var event = doc.createEvent("MouseEvents");
                          event.initMouseEvent(
                              "click", true, false, view, 0, 0, 0, 0, 0
                              , false, false, false, false, 0, null
                          );
                          save_link.dispatchEvent(event);
                          filesaver.readyState = filesaver.DONE;
                          filesaver.savedAs = filesaver.SAVEDASBLOB;
                          dispatch_all();
                          return;
                      }
                      // Object and web filesystem URLs have a problem saving in Google Chrome when
                      // viewed in a tab, so I force save with application/octet-stream
                      // http://code.google.com/p/chromium/issues/detail?id=91158
                      if (view.chrome && type && type !== force_saveable_type) {
                          slice = blob.slice || blob.webkitSlice;
                          blob = slice.call(blob, 0, blob.size, force_saveable_type);
                          blob_changed = true;
                      }
                      // Since I can't be sure that the guessed media type will trigger a download
                      // in WebKit, I append .download to the filename.
                      // https://bugs.webkit.org/show_bug.cgi?id=65440
                      if (webkit_req_fs && name !== "download") {
                          name += ".download";
                      }
                      if (type === force_saveable_type || webkit_req_fs) {
                          target_view = view;
                      }
                      if (!req_fs) {
                          fs_error();
                          return;
                      }
                      fs_min_size += blob.size;
                      req_fs(view.TEMPORARY, fs_min_size, abortable(function(fs) {
                          fs.root.getDirectory("saved", create_if_not_found, abortable(function(dir) {
                              var save = function() {
                                  dir.getFile(name, create_if_not_found, abortable(function(file) {
                                      file.createWriter(abortable(function(writer) {
                                          writer.onwriteend = function(event) {
                                              target_view.location.href = file.toURL();
                                              deletion_queue.push(file);
                                              filesaver.readyState = filesaver.DONE;
                                              filesaver.savedAs = filesaver.SAVEDASBLOB;
                                              dispatch(filesaver, "writeend", event);
                                          };
                                          writer.onerror = function() {
                                              var error = writer.error;
                                              if (error.code !== error.ABORT_ERR) {
                                                  fs_error();
                                              }
                                          };
                                          "writestart progress write abort".split(" ").forEach(function(event) {
                                              writer["on" + event] = filesaver["on" + event];
                                          });
                                          writer.write(blob);
                                          filesaver.abort = function() {
                                              writer.abort();
                                              filesaver.readyState = filesaver.DONE;
                                              filesaver.savedAs = filesaver.FAILED;
                                          };
                                          filesaver.readyState = filesaver.WRITING;
                                      }), fs_error);
                                  }), fs_error);
                              };
                              dir.getFile(name, {create: false}, abortable(function(file) {
                                  // delete file if it already exists
                                  file.remove();
                                  save();
                              }), abortable(function(ex) {
                                  if (ex.code === ex.NOT_FOUND_ERR) {
                                      save();
                                  } else {
                                      fs_error();
                                  }
                              }));
                          }), fs_error);
                      }), fs_error);
                  }
                  , FS_proto = FileSaver.prototype
                  , saveAs = function(blob, name) {
                      return new FileSaver(blob, name);
                  }
                  ;
              FS_proto.abort = function() {
                  var filesaver = this;
                  filesaver.readyState = filesaver.DONE;
                  filesaver.savedAs = filesaver.FAILED;
                  dispatch(filesaver, "abort");
              };
              FS_proto.readyState = FS_proto.INIT = 0;
              FS_proto.WRITING = 1;
              FS_proto.DONE = 2;
              FS_proto.FAILED = -1;
              FS_proto.SAVEDASBLOB = 1;
              FS_proto.SAVEDASURI = 2;
              FS_proto.SAVEDASUNKNOWN = 3;

              FS_proto.error =
                  FS_proto.onwritestart =
                      FS_proto.onprogress =
                          FS_proto.onwrite =
                              FS_proto.onabort =
                                  FS_proto.onerror =
                                      FS_proto.onwriteend =
                                          null;

              view.addEventListener("unload", process_deletion_queue, false);
              saveAs.unload = function() {
                  process_deletion_queue();
                  view.removeEventListener("unload", process_deletion_queue, false);
              };
              return saveAs;
          }(
              typeof self !== "undefined" && self
              || typeof window !== "undefined" && window
              || this.content
          ));
  };

  eje.prototype.saveExcel=function () {
      this.instance(this.changeData(this.data,this.filter));
  };
  window.ExportJsonExcel=eje;
})(window);
