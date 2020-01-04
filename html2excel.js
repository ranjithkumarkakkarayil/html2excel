
var tcdoptionsrowid,
tcdoptionscolid,
startcol,
startrow,
endcol,
endrow,
startsel,
rowcount;

(function($) {
    $.fn.html2excel = function(options) {

        // Variables
        var table,
            totalrows,
            totalcols;

        var table = $(this);

        totalrows = table.find("tbody tr").length;
        totalcols = table.find("thead tr th").length;

        table.addClass("html2excel");

        addHeadings(table, totalrows, totalcols);
		table_pagination(table);
        loadTableOptions(table);
		table_sort(table);
		table_responsive(table);
		table_inline_editing(table);
		right_click(table);
		create_pivot(table);
		create_charts(table);
    };
}(jQuery));

function addHeadings(table, totalrows, totalcols) {
    table.find("thead tr").prepend("<th class='html2excel_first_cell'></th>")
    for (i = 1; i <= totalrows; i++) {
        table.find("tbody tr:nth-child(" + i + ")").prepend("<td class='headings'>" + i + "</td>");
    }
}

function loadTableOptions(table) {

    // Export table
    var export_table = '<a id="tcd_export_excel" data-toggle="tooltip" title="Export to excel"><i class="fas fa-file-excel"></i><br/>Save as Excel</a>' +
		'<a id="tcd_export_csv" data-toggle="tooltip" title="Export to csv"><i class="fas fa-file-csv"></i><br/>Save as CSV</a>' +
        '<a id="tcd_export_json" data-toggle="tooltip" title="Export to json"><i class="fas fa-file-code"></i><br/>Save as JSON</a>' +
        '<a id="tcd_export_copy" data-toggle="tooltip" title="Copy to clipboard"><i class="fas fa-copy"></i><br/>Copy table</a>';

    // Actions
    var actions = '<a id="tcd_table"><i class="fas fa-table"></i><br/>Toggle table</a>' +
        '<a id="tcd_options_paging"><i class="fas fa-ellipsis-h"></i><br/>Toggle paging</a>'+
		'<a id="tcd_options_freeze_pa"><i class="fas fa-border-style"></i><br/>Freeze pane</a>'+
		'<a id="tcd_options_freeze_tr"><i class="fas fa-arrows-alt-h"></i><br/>Freeze header</a>'+
		'<a id="tcd_options_freeze_tc"><i class="fas fa-arrows-alt-v"></i><br/>Freeze 1st col</a>';

    // Data
    var dataandformat = '<a id="tcd_options_pivot"><i class="fas fa-filter"></i><br/>Pivot table</a>' +
		'<a id="tcd_option_graphs"><i class="fas fa-chart-bar"></i><br/>Graphs</a>' +
        '<a id="tcd_option_duplicate"><i class="fas fa-check-double"></i><br/>Duplicates</a>' +
        '<a id="tcd_options_conditional_format"><i class="fas fa-palette"></i><br/>Conditional formatting</a>' +
        '<a id="tcd_option_format"><i class="fas fa-align-center"></i><br/>Format cells</a>';

    // Context menu				
    var rightclick_menu = '<div id="contextMenu" class="dropdown clearfix">' +
        '<ul class="dropdown-menu" role="menu" aria-labelledby="dropdownMenu" style="display:block;position:static;">' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="cut" class="btn btn-link width100 tcd_context_options">Cut</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="copy" class="btn btn-link width100 tcd_context_options">Copy</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="paste" class="btn btn-link width100 tcd_context_options">Paste</a></li>' +
        '<li class="divider"></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="insertcolright" class="btn btn-link width100 tcd_context_options">Insert column right</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="insertcolleft" class="btn btn-link width100 tcd_context_options">Insert column left</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="insertrowtop" class="btn btn-link width100 tcd_context_options">Insert row top</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="insertrowbottom" class="btn btn-link width100 tcd_context_options">Insert row bottom</a></li>' +
        '<li class="divider"></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="deleterow" class="btn btn-link width100 tcd_context_options">Delete row</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="deletecol" class="btn btn-link width100 tcd_context_options">Delete column</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="clearcontent" class="btn btn-link width100 tcd_context_options">Clear contents</a></li>' +
        '<li class="divider"></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="mergecells" class="btn btn-link width100 tcd_context_options">Merge cells</a></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="mergerows" class="btn btn-link width100 tcd_context_options">Merge rows</a></li>' +
        '<li class="divider"></li>' +
        '<li><a tabindex="-1" href="#" data-tcd-context-option="formatcells" class="btn btn-link width100 tcd_context_options">Format cell</a></li>' +
        '</ul>' +
        '</div>';

    var conditional_format = '<div class="modal fade" id="modal_conditional_formating" tabindex="-1" role="dialog" aria-labelledby="modal_conditional_formatingLabel" aria-hidden="true">' +
        '<div class="modal-dialog" role="document">                                                                                                                 ' +
        '  	<div class="modal-content">                                                                                                                          ' +
        '  	  	<div class="modal-body html2excel_modal">                                                                                                                         ' +
        '  	  	  	<p>Conditional formatting: Setup rule</p><hr/>                                                                                               ' +
        '			<table class="width100">                                                                                                                         ' +
        '				<tr>                                                                                                                                         ' +
        '					<td>                                                                                                                                     ' +
        '						<select id="tcd_conditional_format_step1">                                                                                           ' +
        '							<option value="1">Format only cells that contain</option>                                                                        ' +
        '						</select>                                                                                                                            ' +
        '					</td>                                                                                                                                    ' +
        '					<td>                                                                                                                                     ' +
        '						<select id="tcd_conditional_format_step2">                                                                                           ' +
        '							<option value="1">Cell value</option>                                                                                            ' +
        '							<option value="2">Specific text</option>                                                                                         ' +
        '							<option value="3">Blanks</option>                                                                                                ' +
        '							<option value="4">No blanks</option>                                                                                             ' +
        '						</select>                                                                                                                            ' +
        '					</td>                                                                                                                                    ' +
        '					<td id="tcd_td_conditional_format_step3">                                                                                                ' +
        '						<select id="tcd_conditional_format_step3a">                                                                                          ' +
        '							<option value="1">between</option>                                                                                               ' +
        '							<option value="2">not between</option>                                                                                           ' +
        '							<option value="3">equal to</option>                                                                                              ' +
        '							<option value="4">not equal to</option>                                                                                          ' +
        '							<option value="5" selected="selected">greater than</option>                                                                      ' +
        '							<option value="6">less than</option>                                                                                             ' +
        '							<option value="7">greater than or equal to</option>                                                                              ' +
        '							<option value="8">less than or equal to</option>                                                                                 ' +
        '						</select>                                                                                                                            ' +
        '						<select id="tcd_conditional_format_step3b" class="hide">                                                                             ' +
        '							<option value="1">containing</option>                                                                                            ' +
        '							<option value="2">not containing</option>                                                                                        ' +
        '							<option value="3">begining with</option>                                                                                         ' +
        '							<option value="4">ending with</option>                                                                                           ' +
        '						</select>                                                                                                                            ' +
        '					</td>                                                                                                                                    ' +
        '				</tr>                                                                                                                                        ' +
        '				<tr id="tcd_td_conditional_format_tr_value">                                                                                                 ' +
        '					<td><input type="text" id="tcd_conditional_format_value1" class="width100" placeholder="Enter value" /></td>                             ' +
        '					<td class="hide tcd_td_conditional_format_value2 text-center"> and </td>                                                                 ' +
        '					<td class="hide tcd_td_conditional_format_value2"><input type="text" id="tcd_conditional_format_value2" class="width100" placeholder="Enter value" /></td>' +
        '				</tr>                                                                                                                                                         ' +
        '				<tr>                                                                                                                                                          ' +
        '					<td colspan="3">                                                                                                                                          ' +
        '						Set font color <br/>                                                                                                                                  ' +
        '                                                                                                                                                                            ' +
        '						<div class="colorPicker">                                                                                                                             ' +
        '							<input class="black" type="radio" name="font-color" value="black" id="font-color-black" />                                                        ' +
        '							<label class="black" for="font-color-black">black</label>                                                                                         ' +
        '							<input class="white" type="radio" name="font-color" value="white" id="font-color-white" checked="checked"/>                                       ' +
        '							<label class="white" for="font-color-white">white</label>                                                                                         ' +
        '						  	<input class="red" type="radio" name="font-color" value="red" id="font-color-red"/>                                                               ' +
        '						  	<label class="red" for="font-color-red">red</label>                                                                                               ' +
        '						  	<input class="orange" type="radio" name="font-color" value="orange" id="font-color-orange"/>                                                      ' +
        '						  	<label class="orange" for="font-color-orange">orange</label>                                                                                      ' +
        '						  	<input class="yellow" type="radio" name="font-color" value="yellow" id="font-color-yellow"/>                                                      ' +
        '						  	<label class="yellow" for="font-color-yellow">yellow</label>                                                                                      ' +
        '						  	<input class="green" type="radio" name="font-color" value="green" id="font-color-green"/>                                                         ' +
        '						  	<label class="green" for="font-color-green">green</label>                                                                                         ' +
        '						  	<input class="blue" type="radio" name="font-color" value="blue" id="font-color-blue"/>                                                            ' +
        '						  	<label class="blue" for="font-color-blue">blue</label>                                                                                            ' +
        '						  	<input class="indigo" type="radio" name="font-color" value="indigo" id="font-color-indigo"/>                                                      ' +
        '						  	<label class="indigo" for="font-color-indigo">indigo</label>                                                                                      ' +
        '						  	<input class="violet" type="radio" name="font-color" value="violet" id="font-color-violet"/>                                                      ' +
        '						  	<label class="violet" for="font-color-violet">violet</label>                                                                                      ' +
        '						</div>                                                                                                                                                ' +
        '					</td>                                                                                                                                                     ' +
        '				</tr>                                                                                                                                                         ' +
        '				<tr>                                                                                                                                                          ' +
        '					<td colspan="3">                                                                                                                                          ' +
        '						Set cell color <br/>                                                                                                                                  ' +
        '                                                                                                                                                                            ' +
        '						<div class="colorPicker">                                                                                                                             ' +
        '							<input class="black" type="radio" name="cell-color" value="black" id="cell-color-black"/>                                                         ' +
        '						  	<label class="black" for="cell-color-black">black</label>                                                                                         ' +
        '							<input class="white" type="radio" name="cell-color" value="white" id="cell-color-white" />                                                        ' +
        '							<label class="white" for="cell-color-white">white</label>                                                                                         ' +
        '						  	<input class="red" type="radio" name="cell-color" value="red" id="cell-color-red" checked="checked" />                                            ' +
        '						  	<label class="red" for="cell-color-red">red</label>                                                                                               ' +
        '						  	<input class="orange" type="radio" name="cell-color" value="orange" id="cell-color-orange"/>                                                      ' +
        '						  	<label class="orange" for="cell-color-orange">orange</label>                                                                                      ' +
        '						  	<input class="yellow" type="radio" name="cell-color" value="yellow" id="cell-color-yellow"/>                                                      ' +
        '						  	<label class="yellow" for="cell-color-yellow">yellow</label>                                                                                      ' +
        '						  	<input class="green" type="radio" name="cell-color" value="green" id="cell-color-green"/>                                                         ' +
        '						  	<label class="green" for="cell-color-green">green</label>                                                                                         ' +
        '						  	<input class="blue" type="radio" name="cell-color" value="blue" id="cell-color-blue"/>                                                            ' +
        '						  	<label class="blue" for="cell-color-blue">blue</label>                                                                                            ' +
        '						  	<input class="indigo" type="radio" name="cell-color" value="indigo" id="cell-color-indigo"/>                                                      ' +
        '						  	<label class="indigo" for="cell-color-indigo">indigo</label>                                                                                      ' +
        '						  	<input class="violet" type="radio" name="cell-color" value="violet" id="cell-color-violet"/>                                                      ' +
        '						  	<label class="violet" for="cell-color-violet">violet</label>                                                                                      ' +
        '						</div>                                                                                                                                                ' +
        '					</td>                                                                                                                                                     ' +
        '				</tr>                                                                                                                                                         ' +
        '			</table>                                                                                                                                                          ' +
        '			<hr/>                                                                                                                                                             ' +
        '			<a id="tcd_conditional_format_add_rule" class="btn btn-primary btn-xs"><i class="fas fa-plus-circle"></i> Add rule</a>                                            ' +
        '			<hr/>                                                                                                                                                             ' +
        '			<div>                                                                                                                                                             ' +
        '				<table id="tcd_tbl_cond_format_rules"></table>                                                                                                                ' +
        '			</div>                                                                                                                                                            ' +
        '  	  	</div>                                                                                                                                                            ' +
        '  	</div>                                                                                                                                                                ' +
        '</div>' +
        '</div>';


    var duplicates = '<div class="modal fade" id="modal_duplicates" tabindex="-1" role="dialog" aria-labelledby="modal_duplicatesLabel" aria-hidden="true">						 ' +
        '  	<div class="modal-dialog" role="document">                                                                                                   ' +
        '  	  	<div class="modal-content modal-sm">                                                                                                     ' +
        '  	  	  	<div class="modal-body html2excel_modal">                                                                                                             ' +
        '  	  	  	  	<p>Manage duplicates</p><hr />                                                                                                   ' +
        '				<p>                                                                                                                                  ' +
        '					<select id="tcd_dup_action">                                                                                                     ' +
        '						<option value="highlight">Highlight duplicates</option>                                                                      ' +
        '						<option value="remove">Remove duplicates</option>                                                                            ' +
        '					</select>                                                                                                                        ' +
        '				</p>                                                                                                                                 ' +
        '				<p>                                                                                                                                  ' +
        '					<select id="tcd_dup_action_from">                                                                                                ' +
        '						<option value="cells">from selected cells</option>                                                                           ' +
        '						<option value="column">from selected column</option>                                                                         ' +
        '					</select>                                                                                                                        ' +
        '				</p>                                                                                                                                 ' +
        '				<p class="tcd_dup_color">                                                                                                            ' +
        '					<input type="checkbox" id="tcd_dup_ignore_blank" checked="checked" /> Ignore empty cells                                         ' +
        '				</p>                                                                                                                                 ' +
        '				<p class="tcd_dup_color">                                                                                                            ' +
        '					Set font color <br/>                                                                                                             ' +
        '					<div class="colorPicker tcd_dup_color">                                                                                          ' +
        '						<input class="black" type="radio" name="dup-font-color" value="black" id="dup-font-color-black" />                           ' +
        '						<label class="black" for="dup-font-color-black">black</label>                                                                ' +
        '						<input class="white" type="radio" name="dup-font-color" value="white" id="dup-font-color-white" checked="checked"/>          ' +
        '						<label class="white" for="dup-font-color-white">white</label>                                                                ' +
        '					  	<input class="red" type="radio" name="dup-font-color" value="red" id="dup-font-color-red"/>                                  ' +
        '					  	<label class="red" for="dup-font-color-red">red</label>                                                                      ' +
        '					  	<input class="orange" type="radio" name="dup-font-color" value="orange" id="dup-font-color-orange"/>                         ' +
        '					  	<label class="orange" for="dup-font-color-orange">orange</label>                                                             ' +
        '					  	<input class="yellow" type="radio" name="dup-font-color" value="yellow" id="dup-font-color-yellow"/>                         ' +
        '					  	<label class="yellow" for="dup-font-color-yellow">yellow</label>                                                             ' +
        '					  	<input class="green" type="radio" name="dup-font-color" value="green" id="dup-font-color-green"/>                            ' +
        '					  	<label class="green" for="dup-font-color-green">green</label>                                                                ' +
        '					  	<input class="blue" type="radio" name="dup-font-color" value="blue" id="dup-font-color-blue"/>                               ' +
        '					  	<label class="blue" for="dup-font-color-blue">blue</label>                                                                   ' +
        '					  	<input class="indigo" type="radio" name="dup-font-color" value="indigo" id="dup-font-color-indigo"/>                         ' +
        '					  	<label class="indigo" for="dup-font-color-indigo">indigo</label>                                                             ' +
        '					  	<input class="violet" type="radio" name="dup-font-color" value="violet" id="dup-font-color-violet"/>                         ' +
        '					  	<label class="violet" for="dup-font-color-violet">violet</label>                                                             ' +
        '					</div>                                                                                                                           ' +
        '				</p>                                                                                                                                 ' +
        '				<p class="tcd_dup_color">                                                                                                            ' +
        '					Set cell color <br/>                                                                                                             ' +
        '					<div class="colorPicker tcd_dup_color">                                                                                          ' +
        '						<input class="black" type="radio" name="dup-cell-color" value="black" id="dup-cell-color-black"/>                            ' +
        '					  	<label class="black" for="dup-cell-color-black">black</label>                                                                ' +
        '						<input class="white" type="radio" name="dup-cell-color" value="white" id="dup-cell-color-white" />                           ' +
        '						<label class="white" for="dup-cell-color-white">white</label>                                                                ' +
        '					  	<input class="red" type="radio" name="dup-cell-color" value="red" id="dup-cell-color-red" checked="checked" />               ' +
        '					  	<label class="red" for="dup-cell-color-red">red</label>                                                                      ' +
        '					  	<input class="orange" type="radio" name="dup-cell-color" value="orange" id="dup-cell-color-orange"/>                         ' +
        '					  	<label class="orange" for="dup-cell-color-orange">orange</label>                                                             ' +
        '					  	<input class="yellow" type="radio" name="dup-cell-color" value="yellow" id="dup-cell-color-yellow"/>                         ' +
        '					  	<label class="yellow" for="dup-cell-color-yellow">yellow</label>                                                             ' +
        '					  	<input class="green" type="radio" name="dup-cell-color" value="green" id="dup-cell-color-green"/>                            ' +
        '					  	<label class="green" for="dup-cell-color-green">green</label>                                                                ' +
        '					  	<input class="blue" type="radio" name="dup-cell-color" value="blue" id="dup-cell-color-blue"/>                               ' +
        '					  	<label class="blue" for="dup-cell-color-blue">blue</label>                                                                   ' +
        '					  	<input class="indigo" type="radio" name="dup-cell-color" value="indigo" id="dup-cell-color-indigo"/>                         ' +
        '					  	<label class="indigo" for="dup-cell-color-indigo">indigo</label>                                                             ' +
        '					  	<input class="violet" type="radio" name="dup-cell-color" value="violet" id="dup-cell-color-violet"/>                         ' +
        '					  	<label class="violet" for="dup-cell-color-violet">violet</label>                                                             ' +
        '					</div>                                                                                                                           ' +
        '				</p>                                                                                                                                 ' +
        '				<hr/>                                                                                                                                ' +
        '				<p><button id="btn_tcd_duplicate_action" class="btn btn-primary btn-xs">Action</button></p>                                          ' +
        '  	  	  	</div>                                                                                                                               ' +
        '  	  	</div>                                                                                                                                   ' +
        '  	</div>' +
        '</div>';

    var format_cells = '<div class="modal fade" id="modal_format" tabindex="-1" role="dialog" aria-labelledby="modal_formatLabel" aria-hidden="true">							' +
        '  	<div class="modal-dialog" role="document">                                                                                                  ' +
        '  	  	<div class="modal-content modal-sm">                                                                                                    ' +
        '  	  	  	<div class="modal-body html2excel_modal">                                                                                                            ' +
        '				<p>Format cells</p><hr/>                                                                                                            ' +
        '				<table>                                                                                                                             ' +
        '					<tr>                                                                                                                            ' +
        '						<td colspan="3">                                                                                                            ' +
        '							<select id="tcd_format_font_style">                                                                                     ' +
        '								<option value="Times New Roman, Times, serif">Times new roman</option>                                              ' +
        '								<option value="Arial, Helvetica, sans-serif">Arial</option>                                                         ' +
        '								<option value="Georgia, serif">Georgia</option>                                                                     ' +
        '								<option value="Comic Sans MS, cursive, sans-serif">Comic Sans</option>                                              ' +
        '								<option value="Verdana, Geneva, sans-serif">Verdana</option>                                                        ' +
        '							</select>                                                                                                               ' +
        '						</td>                                                                                                                       ' +
        '					</tr>                                                                                                                           ' +
        '					<tr>                                                                                                                            ' +
        '						<td>                                                                                                                        ' +
        '							<select id="tcd_format_font_size">                                                                                      ' +
        '								<option value="6">6</option>                                                                                        ' +
        '								<option value="8">8</option>                                                                                        ' +
        '								<option value="9">9</option>                                                                                        ' +
        '								<option value="10">10</option>                                                                                      ' +
        '								<option value="11">11</option>                                                                                      ' +
        '								<option value="12">12</option>                                                                                      ' +
        '								<option value="14">14</option>                                                                                      ' +
        '								<option value="16">16</option>                                                                                      ' +
        '								<option value="18">18</option>                                                                                      ' +
        '								<option value="20">20</option>                                                                                      ' +
        '								<option value="22">22</option>                                                                                      ' +
        '								<option value="24">24</option>                                                                                      ' +
        '								<option value="26">26</option>                                                                                      ' +
        '								<option value="28">28</option>                                                                                      ' +
        '								<option value="36">36</option>                                                                                      ' +
        '								<option value="48">48</option>                                                                                      ' +
        '								<option value="72">72</option>                                                                                      ' +
        '							</select>                                                                                                               ' +
        '						</td>                                                                                                                       ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="up">A <i class="fas fa-arrow-up"></i></a></td>                    ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="down">A <i class="fas fa-arrow-down"></i></a></td>                ' +
        '					</tr>                                                                                                                           ' +
        '					<tr>                                                                                                                            ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="bold"><i class="fas fa-bold"></i></a></td>                        ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="italic"><i class="fas fa-italic"></i></a></td>                    ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="underline"><i class="fas fa-underline"></i></a></td>              ' +
        '					</tr>                                                                                                                           ' +
        '					<tr>                                                                                                                            ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="left"><i class="fas fa-align-left"></i></a></td>                  ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="center"><i class="fas fa-align-center"></i></a></td>              ' +
        '						<td><a class="tcd_format_options" data-tcd-format-option="right"><i class="fas fa-align-right"></i></a></td>                ' +
        '					</tr>                                                                                                                           ' +
        '				</table>                                                                                                                            ' +
        '				<table>                                                                                                                             ' +
        '					<tr><td>Set font color</td></tr>                                                                                                ' +
        '					<tr>                                                                                                                            ' +
        '						<td>                                                                                                                        ' +
        '							<div class="colorPicker">                                                                                               ' +
        '								<input class="black" type="radio" name="format-font-color" value="black" id="format-font-color-black" checked="checked" />	' +
        '								<label class="black" for="format-font-color-black">B</label>                                                                ' +
        '								<input class="white" type="radio" name="format-font-color" value="white" id="format-font-color-white"/>                     ' +
        '								<label class="white" for="format-font-color-white">W</label>                                                                ' +
        '							  	<input class="red" type="radio" name="format-font-color" value="red" id="format-font-color-red"/>                           ' +
        '							  	<label class="red" for="format-font-color-red">R</label>                                                                    ' +
        '							  	<input class="orange" type="radio" name="format-font-color" value="orange" id="format-font-color-orange"/>                  ' +
        '							  	<label class="orange" for="format-font-color-orange">O</label>                                                              ' +
        '							  	<input class="yellow" type="radio" name="format-font-color" value="yellow" id="format-font-color-yellow"/>                  ' +
        '							  	<label class="yellow" for="format-font-color-yellow">Y</label>                                                              ' +
        '							  	<input class="green" type="radio" name="format-font-color" value="green" id="format-font-color-green"/>                     ' +
        '							  	<label class="green" for="format-font-color-green">G</label>                                                                ' +
        '							  	<input class="blue" type="radio" name="format-font-color" value="blue" id="format-font-color-blue"/>                        ' +
        '							  	<label class="blue" for="format-font-color-blue">B</label>                                                                  ' +
        '							  	<input class="indigo" type="radio" name="format-font-color" value="indigo" id="format-font-color-indigo"/>                  ' +
        '							  	<label class="indigo" for="format-font-color-indigo">I</label>                                                              ' +
        '							  	<input class="violet" type="radio" name="format-font-color" value="violet" id="format-font-color-violet"/>                  ' +
        '							  	<label class="violet" for="format-font-color-violet">V</label>                                                              ' +
        '							</div>                                                                                                                          ' +
        '						</td>                                                                                                                               ' +
        '					</tr>                                                                                                                                   ' +
        '					<tr><td>Set cell color</td></tr>                                                                                                        ' +
        '					<tr>                                                                                                                                    ' +
        '						<td>                                                                                                                                ' +
        '							<div class="colorPicker">                                                                                                       ' +
        '								<input class="black" type="radio" name="format-cell-color" value="black" id="format-cell-color-black" checked="checked" />  ' +
        '								<label class="black" for="format-cell-color-black">B</label>                                                                ' +
        '								<input class="white" type="radio" name="format-cell-color" value="white" id="format-cell-color-white"/>                     ' +
        '								<label class="white" for="format-cell-color-white">W</label>                                                                ' +
        '							  	<input class="red" type="radio" name="format-cell-color" value="red" id="format-cell-color-red"/>                           ' +
        '							  	<label class="red" for="format-cell-color-red">R</label>                                                                    ' +
        '							  	<input class="orange" type="radio" name="format-cell-color" value="orange" id="format-cell-color-orange"/>                  ' +
        '							  	<label class="orange" for="format-cell-color-orange">O</label>                                                              ' +
        '							  	<input class="yellow" type="radio" name="format-cell-color" value="yellow" id="format-cell-color-yellow"/>                  ' +
        '							  	<label class="yellow" for="format-cell-color-yellow">Y</label>                                                              ' +
        '							  	<input class="green" type="radio" name="format-cell-color" value="green" id="format-cell-color-green"/>                     ' +
        '							  	<label class="green" for="format-cell-color-green">G</label>                                                                ' +
        '							  	<input class="blue" type="radio" name="format-cell-color" value="blue" id="format-cell-color-blue"/>                        ' +
        '							  	<label class="blue" for="format-cell-color-blue">B</label>                                                                  ' +
        '							  	<input class="indigo" type="radio" name="format-cell-color" value="indigo" id="format-cell-color-indigo"/>                  ' +
        '							  	<label class="indigo" for="format-cell-color-indigo">I</label>                                                              ' +
        '							  	<input class="violet" type="radio" name="format-cell-color" value="violet" id="format-cell-color-violet"/>                  ' +
        '							  	<label class="violet" for="format-cell-color-violet">V</label>                                                              ' +
        '							</div>                                                                                                                          ' +
        '						</td>                                                                                                                               ' +
        '					</tr>                                                                                                                                   ' +
        '				</table>                                                                                                                                    ' +
        '  	  	  	</div>                                                                                                                                      ' +
        '  	  	</div>                                                                                                                                          ' +
        '  	</div>                                                                                                                                              ' +
        '</div>';

    var table_options = '<div class="row html2excel_action_menu">' +
        '<div class="col-md-2">' +
        '<table>' +
        '<tr><td><a id="action_file" class="show_menu">File</a></td></tr>' +
        '<tr><td><a id="action_action">Actions</a></td></tr>' +
        '<tr><td><a id="action_data">Data & Format</a></td></tr>' +
        '</table></div>' +
        '<div class="col-md-10 html2excel_action_details">' +
        '<div class="action_file show">' + export_table + '</div>' +
        '<div class="action_action hide">' + actions + '</div>' +
        '<div class="action_data hide">' + dataandformat + '</div>' +
        '</div>' +
        '</div>'

	$("body").append(rightclick_menu);
    $("body").append(conditional_format);
    $("body").append(duplicates);
    $("body").append(format_cells);

    table.before(table_options);

	var pivot_chart_containers = '<div id="html2excel_pivot_container"></div><div id="html2excel_chart_container"></div>';
	$(".tc_pager").after(pivot_chart_containers);

	var pivot_container = '<div id="tcd_pivot" class="hide"></div>';
	$("#html2excel_pivot_container").append(pivot_container);
	
	var graph_container = '<div id="div_graph_container" class="hide"></div>';
	$("#html2excel_chart_container").append(graph_container);

    init_table_options(table);
}

function init_table_options(table) {
	var $this = table;
	var table_name = $this.attr("id");
	
    $(".html2excel_action_menu table tr td a").click(function(e) {
        var id = $(this).attr("id");
		$(".html2excel_action_menu table tr td a").removeClass("show_menu");
		$(this).addClass("show_menu");
        $(".html2excel_action_details div").removeClass("show");
        $(".html2excel_action_details div").addClass("hide")
        $("." + id).addClass("show");
    });

    $("#tcd_options_conditional_format").click(function() {
        $("#modal_conditional_formating").modal("show");
    });

    $("#tcd_option_duplicate").click(function() {
        $("#modal_duplicates").modal("show");
    });

    $("#tcd_option_format").click(function() {
        $("#modal_format").modal("show");
    });

	document.getElementById('tcd_conditional_format_step2').addEventListener('change', loadConditionalFormating, false);
    document.getElementById('tcd_conditional_format_step3a').addEventListener('change', loadConditionalFormatingBetween, false);
    document.getElementById('tcd_dup_action').addEventListener('change', dupHighlightChange, false);

    $("#tcd_table").click(function(e) {
        table.toggle();
		$(".tc_pager").toggle();
    });

    $("#tcd_options_paging").click(function(e) {
        table_pagination(table);
    });

    $("#tcd_options_conditional_format").click(function(e) {
        $("#modal_conditional_formating").modal("show");
    });

    $("#tcd_conditional_format_add_rule").click(function(e) {
        add_conditional_formatting_rule(table);
    });

    $("#tcd_options_pivot").click(function(e) {
        $("#tcd_pivot").toggle();
    });

	$("#tcd_option_graphs").click(function(e) {
        $("#div_graph_container").toggle();
    });

    $("#tcd_upload_api").click(function(e) {
        $("#modal_api_loader").modal("show");
    });

    $("#btn_tcd_load_from_api").click(function(e) {
        load_api_to_table();
    });

    $("#tcd_option_duplicate").click(function(e) {
        $("#modal_duplicates").modal("show");
    });

    $("#btn_tcd_duplicate_action").click(function(e) {
        manage_duplicate(table);
    });

    $(".tcd_context_options").click(function(e) {
        manage_context_options($(this).attr("data-tcd-context-option"), table);
    });

	$("#tcd_option_format").click(function(e) {
        $("#modal_format").modal("show");
    });
	
	$(".tcd_format_options").click(function(e) {
        format_options($(this).attr("data-tcd-format-option"), table);
    });

	$("#tcd_format_font_size").change(function(e) {
		format_options("fontselect", table);
	});
	
	$("#tcd_format_font_style").change(function(e) {
		format_options("fontstyle", table);
	});
	
	$("input[name='format-font-color']").change(function(e) {
		format_options("fontcolor", table);
	});
	
	$("input[name='format-cell-color']").change(function(e) {
		format_options("cellcolor", table);
	});	
	
	$("#chart_label").on("change", function(){
	    $("#chart_label").text($(this).val());
	});
	
	$("#tcd_options_freeze_pa").click(function(e) {

    });

	$("#tcd_export_excel").click(function(e) {
		var filename = table_name + ".xls";
        tableToExcel(table_name, 'Sheet 1', filename);
	});

	$("#tcd_export_csv").click(function(e) {
	
	});
	
	$("#tcd_export_json").click(function(e) {
		downloadAsJson($this);
	});
	
	$("#tcd_export_copy").click(function(e) {
		selectElementContents(document.getElementById(table_name));
	});

	$("#tcd_options_freeze_tr").click(function(e) {
		if($this.hasClass("tc_fixed_header"))
			$this.removeClass("tc_fixed_header");
		else
			$this.addClass("tc_fixed_header");
    });

	$("#tcd_options_freeze_tc").click(function(e) {
		if($this.hasClass("tc_fixed_first_col"))
			$this.removeClass("tc_fixed_first_col");
		else
			$this.addClass("tc_fixed_first_col");
    });
	
	
	// Start : Select column
    table.find("thead tr th").click(function(e) {
        var colid = $(this).parent().children().index($(this));
        selected_column = parseInt(colid + 1);
        table.find("tbody tr td:nth-child(" + parseInt(colid + 1) + ")").each(function() {
            if ($(this).css("color") == "rgb(255, 0, 0)")
                $(this).css("color", "black");
            else
                $(this).css("color", "#F00");
        });
    });
    // End : Select column

    // Start : Table cell selection on drag
	table.find("tbody tr td").mousedown(function(e) {
	  if (e.which == 1) { clearselection(table); }
	  $(this).addClass("selected");
	  startcol = $(this).index();
	  startrow = $(this).parent().index() + 1;
	  startsel = true;
	  return false;
	})
	.mouseup(function(e){
	  $(this).addClass("selected");
	  endcol = $(this).closest("td").index();
	  endrow = $(this).closest("td").parent().index() + 1;
	  startsel = false;
	  highightselected(table);
	})
	.mousemove(function(e){
	  if(startsel == true) {
	    $(this).addClass("selected");
	  }
	});
    // End : Table cell selection on drag
}

var isPaginated = true;
function table_pagination($this) {
    $this.each(function() {
        var currentPage = 0;
        var numPerPage = 10;

        var $table = $this;

        if (!$(".tc_pager").length) {
            isPaginated = true;
            var $pager = $('<div class="tc_pager row"></div>');
            var $previous = $('<span class="tc_previous"><<</span>');
            var $next = $('<span class="tc_next">>></span>');
            var $comment = $('<span class="tc_paging_comments"></span>');

            $pager.insertAfter($table).find('span.tc-page-number:first').addClass('active');

            $table.bind('repaginate', function() {
                $table.find('tbody tr').hide();

                $filteredRows = $this.find("tbody tr");

                $filteredRows.slice(currentPage * numPerPage, (currentPage + 1) * numPerPage).show();

                var numRows = $filteredRows.length;
                var numPages = Math.ceil(numRows / numPerPage);

                $pager.find('.tc-page-number, .tc_previous, .tc_next').remove();
                for (var page = 0; page < numPages; page++) {
                    var $newPage = $('<span class="tc-page-number" rel="' + parseInt(page + 1) + '"></span>').text(page + 1).bind('click', {
                        newPage: page
                    }, function(event) {
                        currentPage = event.data['newPage'];
                        $table.trigger('repaginate');
                    })
                    if (page == currentPage) {
                        $newPage.addClass('tc_clickable active');
                    } else {
                        $newPage.addClass('tc_clickable');
                    }
                    $newPage.appendTo($pager)
                }

                $previous.insertBefore('span.tc-page-number:first');
                $next.insertAfter('span.tc-page-number:last');
                $comment.insertAfter('span.tc_next');

                $next.click(function(e) {
                    $previous.addClass('tc_clickable');
                    $pager.find('.active').next('.tc-page-number.tc_clickable').click();
                });
                $previous.click(function(e) {
                    $next.addClass('clickable');
                    $pager.find('.active').prev('.tc-page-number.tc_clickable').click();
                });

                $next.addClass('tc_clickable');
                $previous.addClass('tc_clickable');

                setTimeout(function() {
                    var $active = $pager.find('.tc-page-number.active');
                    if ($active.next('.tc-page-number.tc_clickable').length === 0) {
                        $next.removeClass('tc_clickable');
                    } else if ($active.prev('.tc-page-number.tc_clickable').length === 0) {
                        $previous.removeClass('tc_clickable');
                    }
                });

                var currPage = $pager.find('.tc-page-number.active').attr('rel');
                var startItem = currPage == 1 ? 1 : (((parseInt(currPage) - 1) * parseInt(numPerPage)) + 1);
                var endItem = currPage * numPerPage;
                startItem = startItem || 0;
                endItem = endItem || 0;
                $comment.text("Showing " + startItem + " to " + endItem + " of " + $this.find("tbody tr").length + " entries");
            });
            $table.trigger('repaginate');
        } else {
            $table.find('tbody tr').show();
            $(".tc_pager").remove();
            isPaginated = false;
        }

    });
}

function table_sort($this) {
    $this.find("thead tr th").dblclick(function() {
        var selCol = $(this).parent().children().index($(this));
        var table,
            rows,
            switching,
            i,
            x,
            y,
            shouldSwitch,
            dir,
            switchcount = 0,
            n = selCol;
        table = $this;
        switching = true;
        dir = "asc";
        while (switching) {
            switching = false;
            rows = table.find("tbody tr");
            for (i = 0; i < rows.length - 1; i++) {
                shouldSwitch = false;
                x = rows[i].getElementsByTagName("td")[n];
                y = rows[i + 1].getElementsByTagName("td")[n];
                if (dir == "asc") {
                    if (x.innerHTML.toLowerCase() > y.innerHTML.toLowerCase()) {
                        shouldSwitch = true;
                        break;
                    }
                } else if (dir == "desc") {
                    if (x.innerHTML.toLowerCase() < y.innerHTML.toLowerCase()) {
                        shouldSwitch = true;
                        break;
                    }
                }
            }
            if (shouldSwitch) {
                rows[i].parentNode.insertBefore(rows[i + 1], rows[i]);
                switching = true;
                switchcount++;
            } else {
                if (switchcount == 0 && dir == "asc") {
                    dir = "desc";
                    switching = true;
                }
            }
        }
    });
}

function table_responsive($this) {
    $this.addClass("tc_responsive_table");

    var style = "";
    var total_columns = $this.find("thead tr:first th").length;
    for (i = 1; i < total_columns + 1; i++) {
        var hval = $this.find("thead tr:last-child th:nth-child(" + i + ")").text().replace("↑↓", "");
        style += '.tc_responsive_table td:nth-of-type(' + i + '):before { content: "' + hval + '"; }'
    }
    var tc_responsive_header = style;

    var style = document.createElement('style');
    style.type = 'text/css';
    style.innerHTML = "@media only screen and (max-width: 760px), (min-device-width: 768px) and (max-device-width: 1024px) {" + tc_responsive_header + "}";
    document.getElementsByTagName('head')[0].appendChild(style);
}

function table_inline_editing($this) {
    $this.find("tbody tr td").on('dblclick', function(event) {
        event.preventDefault();

        $(this).attr('contenteditable', 'true');
        $(this).addClass('tc_table_edit_cell');

        if ($(this).text().trim() == "Double click to enter data")
            $(this).text("");

        $(this).focus();
    });

    $this.find("tbody tr td").on('focusout', function(event) {
        event.preventDefault();

        $(this).removeClass('tc_table_edit_cell');
    });
}

function loadConditionalFormating() {
    var type = $("#tcd_conditional_format_step2").val();
    if (type == "1") {
        $("#tcd_td_conditional_format_step3").show();
        $("#tcd_conditional_format_step3a").show();
        $("#tcd_conditional_format_step3b").hide();
        $("#tcd_td_conditional_format_tr_value").show();
    } else if (type == "2") {
        $("#tcd_td_conditional_format_step3").show();
        $("#tcd_conditional_format_step3a").hide();
        $("#tcd_conditional_format_step3b").show();
        $(".tcd_td_conditional_format_value2").hide();
        $("#tcd_td_conditional_format_tr_value").show();
    } else {
        $("#tcd_td_conditional_format_step3").hide();
        $("#tcd_td_conditional_format_tr_value").hide();
    }
}

function loadConditionalFormatingBetween() {
    var type = $("#tcd_conditional_format_step3a").val();
    if (type == "1" || type == "2")
        $(".tcd_td_conditional_format_value2").show();
    else
        $(".tcd_td_conditional_format_value2").hide();
}

function setup_color_picker() {
    var colorList = ['000000', '993300', '333300', '003300', '003366', '000066', '333399', '333333', '660000', 'FF6633', '666633', '336633', '336666', '0066FF', '666699', '666666', 'CC3333', 'FF9933', '99CC33', '669966', '66CCCC', '3366FF', '663366', '999999', 'CC66FF', 'FFCC33', 'FFFF66', '99FF66', '99CCCC', '66CCFF', '993366', 'CCCCCC', 'FF99CC', 'FFCC99', 'FFFF99', 'CCffCC', 'CCFFff', '99CCFF', 'CC99FF', 'FFFFFF'];
    var picker = $('#color-picker');

    for (var i = 0; i < colorList.length; i++) {
        picker.append('<li class="color-item" data-hex="' + '#' + colorList[i] + '" style="background-color:' + '#' + colorList[i] + ';"></li>');
    }

    $('body').click(function() {
        picker.fadeOut();
    });

    $('.call-picker').click(function(event) {
        event.stopPropagation();
        picker.fadeIn();
        picker.children('li').hover(function() {
            var codeHex = $(this).data('hex');
            $('.color-holder').css('background-color', codeHex);
            // alert(codeHex)
        });
    });
}

function add_conditional_formatting_rule(table) {
    var step1a = $("#tcd_conditional_format_step1 option:selected").text();
    var step1av = $("#tcd_conditional_format_step1").val();
    var step1b = $("#tcd_conditional_format_step2 option:selected").text().toLowerCase();
    var step1bv = $("#tcd_conditional_format_step2").val();

    var step1c = "";
    var step1cv = "";
    var val = $("#tcd_conditional_format_step2").val();
    if (val == "1") {
        step1c = $("#tcd_conditional_format_step3a option:selected").text();
        step1cv = $("#tcd_conditional_format_step3a").val();
    } else if (val == "2") {
        step1c = $("#tcd_conditional_format_step3b option:selected").text();
        step1cv = $("#tcd_conditional_format_step3b").val();
    }

    var step1d1 = $("#tcd_conditional_format_value1").val();
    var step1d2 = $("#tcd_conditional_format_value2").val();

    var step1dv = "";
    if (step1d2 == "")
        step1dv = step1d1;
    else
        step1dv = step1d1 + "&" + step1d2;

    var font_color = $("input[name='font-color']:checked").val();
    var cell_color = $("input[name='cell-color']:checked").val();

    var rule = step1a + " > " + step1b + " > " + step1c + " > " + step1dv + " > " + font_color + " > " + cell_color;

    var total_rules = $("#tcd_tbl_cond_format_rules").find("tr").length;
    total_rules = parseInt(total_rules + 1);

    if (step1d1 == "" && (step1bv == "1" || step1bv == "2")) {} else {
        var rule_id = "Rule #" + total_rules;
        var row = "<tr><td>" + rule_id + "</td><td style='background-color:" + cell_color + "; color:" + font_color + "'>" + rule + "</td><td><a class='tcd_con_format_del_rule btn btn-link'><i class='fas fa-times-circle'></i></a></td></tr>";
        $("#tcd_tbl_cond_format_rules").append(row);

        apply_rules(table);

        $('#tcd_tbl_cond_format_rules').on('click', 'tr a.tcd_con_format_del_rule', function(e) {
            e.preventDefault();
            $(this).closest('tr').remove();
            apply_rules(table);
        });
    }

    $("#tcd_conditional_format_value1").val("");
    $("#tcd_conditional_format_value2").val("");
}

function apply_rules(table) {
    reapply_all_rules(table);
    var total_rules = $("#tcd_tbl_cond_format_rules").find("tr").length;
    var $this = $("#tcd_tbl_cond_format_rules");
    for (i = 0; i < total_rules; i++) {
        var rule = $this.find("tr:nth-child(" + parseInt(i + 1) + ") td:nth-child(2)").text();
        var val1 = rule.split('>')[3].split('&')[0];
        var val2 = rule.split('>')[3].split('&')[1];
        var fcolor = rule.split('>')[4];
        var ccolor = rule.split('>')[5];

        switch (rule.split('>')[0].toLowerCase().trim()) {
            case "format only cells that contain":
                switch (rule.split('>')[1].toLowerCase().trim()) {
                    case "cell value":
                        switch (rule.split('>')[2].toLowerCase().trim()) {
                            case "between":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) >= parseFloat(val1) && parseFloat(v) <= parseFloat(val2)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "not between":
                                break;
                            case "equal to":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) == parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "not equal to":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) != parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "greater than":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) > parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "less than":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) < parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "greater than or equal to":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) >= parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                            case "less than or equal to":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!isNaN(v)) {
                                        if (parseFloat(v) <= parseFloat(val1)) {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        }
                                    }
                                });
                                break;
                        }
                        break;
                    case "specific text":
                        switch (rule.split('>')[2].toLowerCase().trim()) {
                            case "containing":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (v.includes(val1.trim())) {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    }
                                });
                                break;
                            case "not containing":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (!v.includes(val1.trim())) {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    }
                                });
                                break;
                            case "begining with":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (v.startsWith(val1.trim())) {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    }
                                });
                                break;
                            case "ending with":
                                table.find("tbody tr td").each(function() {
                                    var v = $(this).text();
                                    if (v.endsWith(val1.trim())) {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    }
                                });
                                break;
                        }
                        break;
                    case "blanks":
                        table.find("tbody tr td").each(function() {
                            var v = $(this).text();
                            if (v == "") {
                                $(this).css("color", fcolor.trim());
                                $(this).css("background-color", ccolor.trim());
                            }
                        });
                        break;
                    case "no blanks":
                        table.find("tbody tr td").each(function() {
                            var v = $(this).text();
                            if (v != "") {
                                $(this).css("color", fcolor.trim());
                                $(this).css("background-color", ccolor.trim());
                            }
                        });
                        break;
                }
                break;
        }
    }
}

function reapply_all_rules(table) {
    table.find("tbody tr td").each(function() {
        $(this).css("color", "black");
        $(this).css("background-color", "white");
    });
}

function clearselection(table) {
  table.find("tr td.selected").removeClass("selected");
}

function highightselected(table) {
	var tempcol = "";
	  if(endcol < startcol) {
	    tempcol = endcol; endcol = startcol; startcol = tempcol;
	  }
	  if(endrow < startrow) {
	    tempcol = endrow; endrow = startrow; startrow = tempcol;
	  }

	  for (var i = startrow; i <= endrow; i++) {
	      var rowCells = table.find("tr").eq(i).find("td");
	      for (var j = startcol; j <= endcol; j++) {
	          rowCells.eq(j).addClass("selected");
	      }
	  }
}

function dupHighlightChange() {
    if ($("#tcd_dup_action").val() == "highlight")
        $(".tcd_dup_color").show();
    else
        $(".tcd_dup_color").hide();
}

function manage_duplicate(table) {
    var fcolor = $("input[name='dup-font-color']:checked").val();
    var ccolor = $("input[name='dup-cell-color']:checked").val();

    var $this = table;

    switch ($("#tcd_dup_action").val()) {
        case "highlight":
            switch ($("#tcd_dup_action_from").val()) {
                case "cells":
                    $this.find("td.selected").each(function() {
                        var txt = $(this).text();
                        $(this).removeClass("selected");
                        $this.find("td.selected").each(function() {
                            if ($(this).text() == txt) {
                                if ($('#tcd_dup_ignore_blank').is(':checked')) {
                                    if ($(this).text() != "") {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    } else {
                                        $(this).css("color", "black");
                                        $(this).css("background-color", "white");
                                    }
                                } else {
                                    $(this).css("color", fcolor.trim());
                                    $(this).css("background-color", ccolor.trim());
                                }
                            }
                        });
                        $(this).addClass("selected");
                    });
                    break;
                case "column":
                    $this.find("td.selected").each(function() {
                        var colid = $(this).parent().children().index($(this)) + 1;
                        $this.find("tr td:nth-child(" + colid + ")").each(function() {
                            var txt = $(this).text();
                            var rowid = $(this).parent().index() + 1;
                            $this.find("tr:not(:nth-child(" + rowid + ")) td:nth-child(" + colid + ")").each(function() {
                                if ($(this).text() == txt) {
                                    if ($('#tcd_dup_ignore_blank').is(':checked')) {
                                        if ($(this).text() != "") {
                                            $(this).css("color", fcolor.trim());
                                            $(this).css("background-color", ccolor.trim());
                                        } else {
                                            $(this).css("color", "black");
                                            $(this).css("background-color", "white");
                                        }
                                    } else {
                                        $(this).css("color", fcolor.trim());
                                        $(this).css("background-color", ccolor.trim());
                                    }
                                }
                            });
                        });
                    });
                    break;
            }
            break;
        case "remove":
            switch ($("#tcd_dup_action_from").val()) {
                case "cells":
                    $this.find("td.selected").each(function() {
                        var txt = $(this).text();
                        $(this).removeClass("selected");
                        $this.find("td.selected").each(function() {
                            if ($(this).text() == txt) {
                                $(this).text("");
                            }
                        });
                        $(this).addClass("selected");
                    });
                    break;
                case "column":
                    $this.find("td.selected").each(function() {
                        var colid = $(this).parent().children().index($(this)) + 1;
                        $this.find("tr td:nth-child(" + colid + ")").each(function() {
                            var txt = $(this).text();
                            var rowid = $(this).parent().index() + 1;
                            $this.find("tr:not(:nth-child(" + rowid + ")) td:nth-child(" + colid + ")").each(function() {
                                if ($(this).text() == txt) {
                                    $(this).text("");
                                }
                            });
                        });
                    });
                    break;
            }
            break;
    }
}

function right_click($this) {
    var $contextMenu = $("#contextMenu");
    $this.find("td").on("contextmenu", function(e) {
        tcdoptionscolid = $(this).parent().children().index($(this)) + 1;
        tcdoptionsrowid = $(this).parent().index() + 1;

        $contextMenu.css({
            display: "block",
            left: e.pageX,
            top: e.pageY
        });
        return false;
    });
    $('html').click(function() {
        $contextMenu.hide();
    });
}

function manage_context_options(option, table) {
    var $this = table;
    let target = $this.find("tbody tr:nth-child(" + tcdoptionsrowid + ") td:nth-child(" + tcdoptionscolid + ")").focus()[0];
	
    switch (option) {
		case "mergerows":
			merge_rows($this);
			break;
		case "mergecells":
			merge_rows($this);
			
			var trows = endrow - startrow + 1;
			var selectedrows = table.find("tr td.selected:last").attr("rowspan");
			if(selectedrows > 1) {
			  trows = parseInt(trows) + parseInt(selectedrows) - 1;
			}
            
			table.find("tr td.selected:first").attr("rowspan",trows);
			table.find("tr td.selected:first").removeClass("selected");
			table.find("tr td.selected").remove();
			
			break;
        case "cut":
            copy_to_clipboard(target);
            $this.find("tbody tr:nth-child(" + tcdoptionsrowid + ") td:nth-child(" + tcdoptionscolid + ")").text("");
            break;
        case "copy":
            copy_to_clipboard(target);
            break;
        case "paste":
            $this.find("tbody tr:nth-child(" + tcdoptionsrowid + ") td:nth-child(" + tcdoptionscolid + ")").text(get_from_clipboard());
            break;
        case "insertcolright":
            var newColumn = [], 
			      colsCount = table.find('tr > td:last').index();
			table.find("tr").each( function(rowIndex) {
			  	var cell = $("<t"+(rowIndex == 0 ?  "h" : "d")+"/>").text((rowIndex == 0 ?  "New col" : ""));
			  	newColumn.push( 
			  	  tcdoptionscolid > colsCount
			  	        ? cell.appendTo(this)
			  	        : cell.insertBefore( $(this).children().eq(tcdoptionscolid) )
			  	);
			});
			
			right_click($this);
            table_inline_editing($this);
            break;
        case "insertcolleft":
			var newColumn = [], 
			      colsCount = table.find('tr > td:last').index();
			table.find("tr").each( function(rowIndex) {
			  	var cell = $("<t"+(rowIndex == 0 ?  "h" : "d")+"/>").text((rowIndex == 0 ?  "New col" : ""));
			  	newColumn.push( 
			  	  (tcdoptionscolid - 1) > colsCount
			  	        ? cell.appendTo(this)
			  	        : cell.insertBefore( $(this).children().eq((tcdoptionscolid - 1)) )
			  	);
			});

			right_click($this);
            table_inline_editing($this);
            break;
        case "insertrowtop":
            var total_columns = $this.find("thead tr:first th").length;
            var html = "";
            for (i = 0; i < total_columns; i++) {
                if (i == 0)
                    html += "<td>Double click to enter data</td>";
                else
                    html += "<td></td>";
            }
            html = "<tr>" + html + "</tr>";
            if (tcdoptionsrowid == 1)
                $this.find('tbody > tr').eq(tcdoptionsrowid - 1).before(html);
            else
                $this.find('tbody > tr').eq(tcdoptionsrowid - 2).after(html);

            right_click($this);
            table_inline_editing($this);
            break;
        case "insertrowbottom":
            var total_columns = $this.find("thead tr:first th").length;
            var html = "";
            for (i = 0; i < total_columns; i++) {
                if (i == 0)
                    html += "<td>Double click to enter data</td>";
                else
                    html += "<td></td>";
            }
            html = "<tr>" + html + "</tr>";
            $this.find('tbody > tr').eq(tcdoptionsrowid - 1).after(html);

            right_click($this);
            table_inline_editing($this);
            break;
        case "deleterow":
            $this.find("tbody tr:nth-child(" + tcdoptionsrowid + ")").remove();
            break;
        case "deletecol":
            $this.find("tbody tr td:nth-child(" + parseInt(tcdoptionscolid) + ")").remove();
            $this.find("thead tr th:nth-child(" + parseInt(tcdoptionscolid) + ")").remove();
            break;
        case "clearcontent":
            $this.find("tbody tr:nth-child(" + tcdoptionsrowid + ") td:nth-child(" + tcdoptionscolid + ")").text("");
            break;
		case "formatcells":
			$("#modal_format").modal("show");
			break;
    }
}

function merge_rows($this) {
	$this.find("tr").each(function(e) {
	        var colspanval = 0;
	        var celllength = $(this).find('td.selected').length;
	        var selectedcells = table.find("tr td.selected:last").attr("colspan");
	        if(selectedcells > 1) {
	          celllength = parseInt(celllength) + parseInt(selectedcells) - 1;
	        }
	        $(this).find("td.selected").each(function(cellindex) {
	            if (cellindex == 0) {
	                if ($(this).attr('colspan') > 1) {
	                    colspanval = $(this).attr('colspan');
	                    celllength = parseInt(celllength) + parseInt(colspanval) - 1;
	                }
	                $(this).attr("colspan", celllength);
	            } else {
	                $(this).remove();
	            }
	        });
	    });
}

function copy_to_clipboard(target) {
	target = $("#tcd_container").find("table tr td.selected")[0];
	
	// var srow = target.find("tr td.selected:first").parent().index();
	// var lrow = target.find("tr td.selected:last").parent().index();
	// var scol = target.find("tr td.selected:first").index();
	// var lcol = target.find("tr td.selected:last").index();
	// var nrows = lrow - srow;
	// var ncols = lcol - scol;

    let range = document.createRange();
    range.selectNodeContents(target);
    let sel = document.getSelection();
    sel.removeAllRanges();
    sel.addRange(range);
    document.execCommand('copy');
}

function get_from_clipboard() {
    navigator.clipboard.readText()
        .then(text => {
            return text;
        });
}

function table_resize($this) {
	var thHeight = $this.find("th:first").height();
	  $this.find("th").resizable({
	      handles: "e",
	      minHeight: thHeight,
	      maxHeight: thHeight,
	      minWidth: 40,
	      resize: function (event, ui) {
	        var sizerID = "#" + $(event.target).attr("id") + "-sizer";
	        $(sizerID).width(ui.size.width);
	      }
	  });
}

function format_options(option, table) {
	var $this = table;
	switch(option){
		case "up":		
			$this.find("td.selected").each(function(){
				var size = $(this).css('font-size');
				size = parseInt(size.substring(0, size.length - 2)) + 1;
				$(this).css('font-size', size + "px");
			});
		break;
		case "down":
			$this.find("td.selected").each(function(){
				var size = $(this).css('font-size');
				size = parseInt(size.substring(0, size.length - 2)) - 1;
				$(this).css('font-size', size + "px");
			});
		break;
		case "bold":
			$this.find("td.selected").each(function(){
				if($(this).hasClass("bold"))
					$(this).removeClass("bold");
				else
					$(this).addClass("bold");
			});
		break;
		case "italic":
			$this.find("td.selected").each(function(){
				if($(this).hasClass("italic"))
					$(this).removeClass("italic");
				else
					$(this).addClass("italic");
			});
		break;
		case "underline":
			$this.find("td.selected").each(function(){
				if($(this).hasClass("underline"))
					$(this).removeClass("underline");
				else
					$(this).addClass("underline");
			});
		break;
		case "left":
			$this.find("td.selected").each(function(){
				$(this).removeClass("left");
				$(this).removeClass("right");
				$(this).removeClass("center");
				
				if($(this).hasClass("left"))
					$(this).removeClass("left");
				else
					$(this).addClass("left");
			});
		break;
		case "center":
			$this.find("td.selected").each(function(){
				$(this).removeClass("left");
				$(this).removeClass("right");
				$(this).removeClass("center");
				
				if($(this).hasClass("center"))
					$(this).removeClass("center");
				else
					$(this).addClass("center");
			});
		break;
		case "right":
			$this.find("td.selected").each(function(){
				$(this).removeClass("left");
				$(this).removeClass("right");
				$(this).removeClass("center");
				
				if($(this).hasClass("right"))
					$(this).removeClass("right");
				else
					$(this).addClass("right");
			});
		break;
		case "fontselect":
			var fontsize = $("#tcd_format_font_size").val();
			$this.find("td.selected").each(function(){
				$(this).css('font-size', fontsize + "px");
			});
		break;
		case "fontstyle":
			var fontstyle = $("#tcd_format_font_style").val();
			$this.find("td.selected").each(function(){
				$(this).css('font-family', fontstyle);
			});
		break;
		case "fontcolor":
			var fontcolor = $("input[name='format-font-color']:checked").val();
			$this.find("td.selected").each(function(){
				$(this).css('color', fontcolor);
			});
		break;
		case "cellcolor":
			var cellcolor = $("input[name='format-cell-color']:checked").val();
			$this.find("td.selected").each(function(){
				$(this).css('background-color', cellcolor);
			});
		break;
		case "rotate":
			var degree = $("#tcd_format_font_rotate").val();
			$this.find("td.selected").each(function(){
				$(this).css('height', $(this).width());
				$(this).css('-moz-transform', 'rotate('+degree+'deg)');
				$(this).css('-webkit-transform', 'rotate('+degree+'deg)');
				$(this).css('-o-transform', 'rotate('+degree+'deg)');
				$(this).css('transform', 'rotate('+degree+'deg)');
			});
		break;
	}
}

// PIVOT TABLE 2.0
// START

function create_pivot(table) {
	setup_pivot_area();
  	
  	var trows = table.find("tbody tr").length;
  	var tcols = table.find("thead tr th").length;
  	
  	var headerslist = [];
  	for(i=1;i<=tcols;i++) {
  	  var header_value = table.find("thead tr th:nth-child("+i+")").text();
  	  headerslist.push(header_value);
  	  $(".dimension_field_name").append("<a id='"+i+"'>"+header_value+"</a>");
  	}
  
	enable_drag_drop_pivot(table);
  	enable_filter_selection(table);
}

function setup_pivot_area() {
	var pivot_container = '<div id="div_pivot_container"></div>';
	$("#tcd_pivot").append(pivot_container);
	
  	var container = '<div class="row"><div class="col-md-3 pivot_dimensions"></div><div class="col-md-9 pivot_values"></div></div>';
  	$("#div_pivot_container").append(container);
  	var dimensions = '<div class="row"><div class="col-md-12"><p>Single level pivot table builder</p><p>Field name</p><div class="dimension_field_name"></div></div></div><div class="row"><div class="col-md-6"><p>Filters</p><div class="dimension_filters dimension_droppable"></div></div><div class="col-md-6"><p>Columns</p><div class="dimension_columns dimension_droppable"></div></div></div><div class="row"><div class="col-md-6"><p>Rows</p><div class="dimension_rows dimension_droppable"></div></div><div class="col-md-6"><p>Values</p><div class="dimension_values dimension_droppable"></div></div></div>';
  	$(".pivot_dimensions").append(dimensions);
  	var actions = '<select id="pivot_value_action" style="display:none"><option value="count">Count of value</option><option value="sum">Sum of value</option><option value="average">Average of value</option><option value="max">Max of value</option><option value="min">Min of value</option></select>';
  	var pivottable_actions = '<div class="row"><div class="col-md-2">'+actions+'</div></div>';
  	$(".pivot_values").append(pivottable_actions);
  	var pivot_table_container = '<div class="pivot_table_container"></div>';
  	$(".pivot_values").append(pivot_table_container);
}

function enable_drag_drop_pivot(table) {
	$('.dimension_field_name a').draggable({
	    cancel: "a.ui-icon",
        revert: "invalid",
        containment: "document",
        helper: "clone",
        cursor: "move",
		stack: ".dimension_field_name a"
	});
	
	$(".dimension_field_name").droppable({
        accept: ".dimension_droppable > a",
        drop: function(event, ui) {
            $(this).append(ui.draggable);
			load_pivot_template(table);
        }
    });
	
	$(".dimension_droppable").droppable({
        accept: ".dimension_field_name > a",
        drop: function(event, ui) {
            $(this).append(ui.draggable.clone());
			load_pivot_template(table);
			
            $(this).find("a").dblclick(function() {
                var elid = ui.draggable.attr("id");
                var container = $(this).parent().attr('class').split(' ')[0]
                $("." + container).find("#" + elid).remove();
                load_pivot_template(table);
            });
			
        }
    });
}

function load_pivot_template(table) {
	// clear elements before recreating them
	$("#pivot_table").remove();
	$("#rowlabelfilter").empty();
	//$("#filter_container").remove();
	
	// pivot math operation
	var pivot_value_action = $("#pivot_value_action").val();
	
	// dimension elements
	var filter_el = 0,
		col_el = 0,
		row_el = 0,
		value_el = 0;
	
	// all data's
	var filterdata_all = [],
		columndata_all = [],
		rowdata_all = [],
		valuedata_all = [],
		uniqueNames_all = [],
		rowlabelfilter_all = [];
	
	// unique data's
	var filterdata = [],
		columndata = [],
		rowdata = [],
		valuedata = [],
		uniqueNames = [],
		rowlabelfilter = [];
		
	// total row and column size for pivot table
	var pivot_col_length = 0,
		pivot_row_length = 0,
		table_row_length = 0;
		
	table_row_length = table.find("tbody tr").length;
	
	// get first dimension element
	filter_el = $(".dimension_filters").find("a:first").attr("id");
	col_el = $(".dimension_columns").find("a:first").attr("id");
	row_el = $(".dimension_rows").find("a:first").attr("id");
	value_el = $(".dimension_values").find("a:first").attr("id");
		
	// push dimension element data to respective arrays
	if(filter_el > 0) {
		table.find("tbody td:nth-child("+filter_el+")").each(function() {
		    if ($.inArray($(this).text(), filterdata) == -1) {
		        filterdata.push($(this).text()); // array to store unique data
				filterdata_all.push($(this).text()); // array to store all data
			}
		});
		
		// filtering out unique. will be used in populating table
		uniqueNames = [];
		$.each(filterdata, function(i, el) {
	        if ($.inArray(el, uniqueNames) === -1) uniqueNames.push(el);
	    });
	    filterdata = uniqueNames;
	
		// sorting the array
		filterdata.sort();
	}
	
	if(col_el > 0) {
		table.find("tbody td:nth-child("+col_el+")").each(function() {
		    if ($.inArray($(this).text(), columndata) == -1) {
		        columndata.push($(this).text());
				columndata_all.push($(this).text());
			}
		});
		
		uniqueNames = [];
		$.each(columndata, function(i, el) {
	        if ($.inArray(el, uniqueNames) === -1) uniqueNames.push(el);
	    });
	
	    columndata = uniqueNames;
	
		columndata.sort();
		
		pivot_col_length = parseInt(columndata.length) + 2;
	}
		
	if(row_el > 0) {
		table.find("tbody td:nth-child("+row_el+")").each(function() {
		    if ($.inArray($(this).text(), rowdata) == -1) {
		        rowdata.push($(this).text());
				rowdata_all.push($(this).text());
			}
		});
		
		uniqueNames = [];
		$.each(rowdata, function(i, el) {
	        if ($.inArray(el, uniqueNames) === -1) uniqueNames.push(el);
	    });
	    rowdata = uniqueNames;
	
		rowdata.sort();
		
		pivot_row_length = parseInt(rowdata.length) + 2;		
	}
		
	if(value_el > 0) {
		$("#pivot_value_action").show();
		$("#pivot_value_action").val("count");
	
		table.find("tbody td:nth-child("+value_el+")").each(function() {
		    if ($.inArray($(this).text(), valuedata) == -1) {
		        valuedata.push($(this).text());
				valuedata_all.push($(this).text());
			}
		});
	}
	else {
		$("#pivot_value_action").hide();
	}	
		

	var pivot_table = '<table id="pivot_table"></table>';
	$(".pivot_table_container").append(pivot_table);
	
	// setup dimension filter dropdown
	if(filter_el > 0) {
		if($(".pivot_table_container").find("#filter_container").length == 0)
		{
			var filter_container_table = '<table id="filter_container"><tr><td colspan="2">Filter by</td></tr><tr><td><span id="span_filter_container"></span></td><td><select id="select_filter_container"></select></td></tr></table>';
			$("#pivot_table").before(filter_container_table);
			
			$("#span_filter_container").text($(".dimension_filters").find("a:first").text());
			
			$.each(filterdata, function(key, value) {   
			     $('#select_filter_container')
			         .append($("<option></option>")
			                    .attr("value",key)
			                    .text(value)); 
			});
			$('#select_filter_container').prepend($("<option value='-1' selected='selected'>All</option>"));
		}
	}
	
	var colFilter = "";
	
	var rgt = [],
		cgt = [];
	
	for(i = 0; i < pivot_row_length; i++) {
		var tr = "";
		var td = "";
		
		for(j = 0; j < pivot_col_length; j++) {
			if (i == 0) {
				if(j == 0) {
					td += "<th><select id='rowlabelfilter'></select></th>";
				}
				else if(j == pivot_col_length - 1) {
					td += "<th>Grand total</th>";
				}
				else {
					td += "<th>" + columndata[j - 1] + "</th>";
				}
				
				if(j == 1) {
					colFilter += "<th><select id='collabelfilter'></select></th>";
				}
				else {
					colFilter += "<th></th>";
				}
			}
			else if (i == pivot_row_length - 1) {
				if(j == 0) {
					td += "<td>Grand total</td>";
				}
				else if(j == pivot_col_length) {
					td += "<td></td>";
				}
				else {
					td += '<td data-attr-count="" data-attr-sum="" data-attr-avg="" data-attr-max="" data-attr-min=""></td>'; // CGT
				}
			}
			else {
				if(j == 0) {
					td += "<td>" + rowdata[i - 1] + "</td>";
				}
				else if(j == pivot_col_length - 1) {					
					var vcount = rgt.length;
					
					var vsum = 0;
					$.each(rgt, function (index, value) {
					    vsum = parseInt(vsum) + parseInt(value);
					});
					if(vsum == 0 || rgt.length == 0)
						var vavg = 0;
					else
						var vavg = (parseInt(vsum) / parseInt(rgt.length)).toFixed(2)
						
					var vmax = Math.max.apply(null, rgt) == -Infinity ? "" : Math.max.apply(null, rgt);
					vmax == 0 ? "" : vmax;
					var vmin = Math.min.apply(null, rgt) == Infinity ? "" : Math.min.apply(null, rgt);
					vmin == 0 ? "" : vmin;
					
					td += '<td data-attr-count="'+vcount+'" data-attr-sum="'+vsum+'" data-attr-avg="'+vavg+'" data-attr-max="'+vmax+'" data-attr-min="'+vmin+'">' + vcount +'</td>';
										
					rgt = [];
				}
				else {
					var value = "";
					
					if(value_el > 0) {
						
						var rdata = rowdata[i - 1];
						var cdata = columndata[j - 1];
						var vdata = valuedata_all[j - 1];
						
						var count = 0;
						var sum = 0;
						
						var ftemp = [];
						var fdata = "";
						if(filter_el > 0) {
                        	table.find("tbody tr td:nth-child(" + filter_el + ")").each(function(index) {
                        	    ftemp.push($(this).text());
                        	});
							fdata = $( "#select_filter_container option:selected" ).text();
						}
						
						var rtemp = [];
                        table.find("tbody tr td:nth-child(" + row_el + ")").each(function(index) {
                            rtemp.push($(this).text());
                        });

                        var ctemp = [];
                        table.find("tbody tr td:nth-child(" + col_el + ")").each(function(index) {
                            ctemp.push($(this).text());
                        });

                        var vtemp = [];
                        table.find("tbody tr td:nth-child(" + value_el + ")").each(function(index) {
                            var v = $(this).text().replace(/[^a-zA-Z0-9]/g, "")
                            vtemp.push(v);
                        });
						
						var minmax = [];
						for(x=0; x<rtemp.length; x++) {
							if(filter_el == 0 || fdata == "All" || fdata == "") {
								if ((rtemp[x] == rdata) && (ctemp[x] == cdata) && (vtemp[x].length > 0)) {
									count++;
									sum += parseInt(vtemp[x]);
									minmax.push(vtemp[x]);
									rgt.push(vtemp[x]);
								}
							}
							else {
								if ((rtemp[x] == rdata) && (ctemp[x] == cdata) && (vtemp[x].length > 0) && (ftemp[x] == fdata)) {
									count++;
									sum += parseInt(vtemp[x]);
									minmax.push(vtemp[x]);
									rgt.push(vtemp[x]);
								}
							}
						}
						
						var vcount = count;
						var vsum = sum;
						if(sum == 0 ||  count == 0)
							var vavg = 0;
						else
							var vavg = (sum / count).toFixed(2);
						var vmax = Math.max.apply(null, minmax) == -Infinity ? 0 : Math.max.apply(null, minmax);
						vmax == 0 ? "" : vmax;
						var vmin = Math.min.apply(null, minmax) == Infinity ? 0 : Math.min.apply(null, minmax);
						vmin == 0 ? "" : vmin;
						
						value = vcount;
					}
										
					if(value == 0)
						td += "<td></td>";
					else
						td += '<td data-attr-count="'+vcount+'" data-attr-sum="'+vsum+'" data-attr-avg="'+vavg+'" data-attr-max="'+vmax+'" data-attr-min="'+vmin+'">' + value +'</td>';
				}
			}
		}
		
		if(i == 0) {
			tr = "<thead><tr>" + td + "</tr></thead>";
		}
		else if(i == 1) {
			tr = "<tbody><tr>" + td + "</tr>";
		}
		else if(i == pivot_row_length) {
			tr = "<tr>" + td + "</tr></tbody>";
		}
		else {
			tr = "<tr>" + td + "</tr>";
		}
		
		$("#pivot_table").append(tr);
	}
	
	$("#pivot_table").find("thead").prepend("<tr>" + colFilter + "</tr>");
	
	
	if(value_el > 0) {
		var trows = $("#pivot_table tbody tr").length;
		var tcols = $("#pivot_table thead tr:last-child th").length;
		
		for(x = 1; x < tcols; x++) {
			var cgt_count_a = [];
			var cgt_sum_a = [];
			var cgt_avg_a = [];
			var cgt_max_a = [];
			var cgt_min_a = [];
			
			$("#pivot_table tbody tr").each(function() {
				$(this).find("td").each(function(eqindex){
					if(eqindex == x) {
						if($(this).text() != 0) {
							cgt_count_a.push($(this).attr("data-attr-count"));
							cgt_sum_a.push($(this).attr("data-attr-sum"));
							cgt_avg_a.push($(this).attr("data-attr-avg"));
							cgt_max_a.push($(this).attr("data-attr-max"));
							cgt_min_a.push($(this).attr("data-attr-min"));
						}
					}
				})
			});
			
			var cgt_count = 0;
			$.each(cgt_count_a, function (index, value) {
				if(!isNaN(value) && value.length != 0) {
			    	cgt_count = parseInt(cgt_count) + parseInt(value);
				}
			});
			
			var cgt_sum = 0;
			$.each(cgt_sum_a, function (index, value) {
				if(!isNaN(value) && value.length != 0) {
			    	cgt_sum = parseInt(cgt_sum) + parseInt(value);
				}
			});
			
			
			var cgt_avg_sum = 0;
			var cgt_avg_count = 0;
			$.each(cgt_avg_a, function (index, value) {
				if(!isNaN(value) && value.length != 0) {
			    	cgt_avg_sum = parseInt(cgt_avg_sum) + parseInt(value);
					if(value != 0)
						cgt_avg_count++;
				}
			});
			var cgt_avg = (cgt_avg_sum / cgt_avg_count).toFixed(2);
				
			var cgt_max = Math.max.apply(null, cgt_max_a) == -Infinity ? "" : Math.max.apply(null, cgt_max_a);
			cgt_max == 0 ? "" : cgt_max;
			var cgt_min = Math.min.apply(null, cgt_min_a) == Infinity ? "" : Math.min.apply(null, cgt_min_a);
			cgt_min == 0 ? "" : cgt_min;
			
			cgt = [];
			
			$("#pivot_table tbody tr:last-child td").eq(x).attr("data-attr-count", cgt_count);
			$("#pivot_table tbody tr:last-child td").eq(x).attr("data-attr-sum", cgt_sum);
			$("#pivot_table tbody tr:last-child td").eq(x).attr("data-attr-avg", cgt_avg);
			$("#pivot_table tbody tr:last-child td").eq(x).attr("data-attr-max", cgt_max);
			$("#pivot_table tbody tr:last-child td").eq(x).attr("data-attr-min", cgt_min);
			
			$("#pivot_table tbody tr:last-child td").eq(x).text(cgt_count);
		}
	}
	
	
	$.each(rowdata, function(key, value) {   
	     $('#rowlabelfilter')
	         .append($("<option></option>")
	                    .attr("value",key)
	                    .text(value)); 
	});	
	$('#rowlabelfilter').prepend($("<option value='-1' selected='selected'>Row labels</option>"));
	
	$.each(columndata, function(key, value) {   
	     $('#collabelfilter')
	         .append($("<option></option>")
	                    .attr("value",key)
	                    .text(value)); 
	});
	$('#collabelfilter').prepend($("<option value='-1' selected='selected'>Column labels</option>"));	
		
	enable_filter_selection(table);
}

function enable_filter_selection(table) {
	var index = -1;
	
	$("#pivot_value_action").change(function(){
		//load_pivot_template(table);
		switch($(this).val()) {
			case "count":
				$("#pivot_table tbody tr td").each(function(){
					var attr = $(this).attr('data-attr-count');
					if (typeof attr !== typeof undefined && attr !== false) {
					    $(this).text(attr);
					}
				});
			break;
			case "sum":
				$("#pivot_table tbody tr td").each(function(){
					var attr = $(this).attr('data-attr-sum');
					if (typeof attr !== typeof undefined && attr !== false) {
					    $(this).text(attr);
					}
				});
			break;
			case "average":
				$("#pivot_table tbody tr td").each(function(){
					var attr = $(this).attr('data-attr-avg');
					if (typeof attr !== typeof undefined && attr !== false) {
					    $(this).text(attr);
					}
				});
			break;
			case "min":
				$("#pivot_table tbody tr td").each(function(){
					var attr = $(this).attr('data-attr-min');
					if (typeof attr !== typeof undefined && attr !== false) {
					    $(this).text(attr);
					}
				});
			break;
			case "max":
				$("#pivot_table tbody tr td").each(function(){
					var attr = $(this).attr('data-attr-max');
					if (typeof attr !== typeof undefined && attr !== false) {
					    $(this).text(attr);
					}
				});
			break;
		}
	});
	
	$("#rowlabelfilter").change(function() {
		var text = $("#rowlabelfilter option:selected").text();
		
		if(text == "Row labels") {
			$("#pivot_table").find("tbody tr").css("color", "black");
			$("#pivot_table").find("tbody tr").css("background-color", "transparent");
		}
		else {
			$("#pivot_table").find("tbody tr").each(function(){
				$(this).find("td").eq(0).each(function(cindex){
					if($(this).text() == text) {
						index = $(this).parent().index();
					}
				});
			});
			
			$("#pivot_table").find("tbody tr").css("color", "black");
			$("#pivot_table").find("tbody tr").css("background-color", "transparent");
			$("#pivot_table").find("tbody tr").eq(index).css("color", "white");
			$("#pivot_table").find("tbody tr").eq(index).css("background-color", "#D50000");
		}
    });

	$("#collabelfilter").change(function() {
		var text = $("#collabelfilter option:selected").text();
		
		if(text == "Column labels") {
			$("td, th").css("color", "black");
			$("td, th").css("background-color", "transparent");
		}
		else {
			$("#pivot_table").find("thead tr").each(function(){
				$(this).find("th").each(function(){
					if($(this).text() == text) {
						index = $(this).index();
					}
				}) 
			});
			
			$("#pivot_table").find("td, th").css("color", "black");
			$("#pivot_table").find("td, th").css("background-color", "transparent");
			$("#pivot_table").find("td, th").filter(":nth-child(" + (index + 1) + ")").css("background-color", "#D50000");
			$("#pivot_table").find("td, th").filter(":nth-child(" + (index + 1) + ")").css("color", "white");
		}
    });

	$("#select_filter_container").change(function(){
		load_pivot_template(table);
	});
}

// END
// PIVOT TABLE 2.0


// GRAPHS 2.0
// START

function create_charts(table) {
    setup_graph_area(table);
    enable_drag_drop_pivot(table);
}

function setup_graph_area(table) {
    var container = '<div class="row"><div class="col-md-2 graph_dimensions"></div><div class="col-md-8 graph_area"></div><div class="col-md-2 graph_type"></div></div>';
    $("#div_graph_container").append(container);

    var dimension = '<div><p>Chart dimensions</p><div class="dimension_container"></div></div>';
    $(".graph_dimensions").append(dimension);

    var elements = '<p>Plot chart</p><table class="tbl_chart_elements">' +
        '<tr><td>Chart label</td><td><input type="text" id="txt_chart_label" placeholder="Enter chart label" /></td></tr>' +
        '<tr><td>Row (X-Axis)</td><td><div class="chart_row chart_rc"></div></td></tr>' +
        '<tr><td>Column (Y-Axis)</td><td><div class="chart_column chart_rc"></div></td></tr>' +
        '</table>';
    $(".graph_area").append(elements);

    var graph = '<div class="graph_container"><p><span id="span_chart_label"></span></p><div class="graph_container_graphs"></div></div>';
    $(".graph_area").append(graph);

    var charttype = '<p>Select chart type</p>'
			+'<a id="bar"class="chart_active">Bar chart</a>'
			+'<a id="column">Column chart</a>'
			+'<a id="pie">Pie chart</a>'
			+'<a id="doughnut">Doughnut chart</a>'
			+'<a id="line">Line chart</a>';
			//+'<a id="area" class="chart_active">Area chart</a>';
			
    $(".graph_type").append(charttype);

    var trows = table.find("tbody tr").length;
    var tcols = table.find("thead tr th").length;

    for (i = 1; i <= tcols; i++) {
        var header_value = table.find("thead tr th:nth-child(" + i + ")").text();
        $(".dimension_container").append("<a id='" + i + "'>" + header_value + "</a>");
    }

    $("#txt_chart_label").on('input', function() {
        $("#span_chart_label").text($(this).val())
    });

	$(".graph_type a").on("click", function(){
		$('.graph_type').children('a').each(function () {
			$(this).removeClass("chart_active");
		});
		$(this).addClass("chart_active");
		
	    var row_el = $(".chart_row").find("a:first").attr("id");
	    var col_el = $(".chart_column").find("a:first").attr("id");
		if (row_el > 0 && col_el > 0) {
			load_chart(table, get_chart_type());
		}
	});
}

function enable_drag_drop_pivot(table) {
    $('.dimension_container a').draggable({
        cancel: "a.ui-icon",
        revert: "invalid",
        containment: "document",
        helper: "clone",
        cursor: "move",
        stack: ".dimension_container a"
    });

    $(".dimension_container").droppable({
        accept: ".chart_rc > a",
        drop: function(event, ui) {
            $(this).append(ui.draggable);

            load_chart(table, get_chart_type());
        }
    });

    $(".chart_rc").droppable({
        accept: ".dimension_container > a",
        drop: function(event, ui) {
            $(this).append(ui.draggable.clone());

            load_chart(table, get_chart_type());

            $(this).find("a").dblclick(function() {
                var elid = ui.draggable.attr("id");
                var container = $(this).parent().attr('class').split(' ')[0]
                $("." + container).find("#" + elid).remove();
                load_chart(table, get_chart_type());
            });
        }
    });
}

function get_chart_type() {
	var charttype = "";
	$('.graph_type').children('a').each(function () {
		if(this.className == "chart_active")
			charttype = this.id;
	});
	return charttype;
}

function load_chart(table, charttype) {

    $(".graph_container_graphs").empty();

    var row_el = 0,
        col_el = 0,
        columndata = [],
        columndata_all = [],
        rowdata = [],
        rowdata_all = [],
        uniqueNames = [];

    // get first dimension element
    row_el = $(".chart_row").find("a:first").attr("id");
    col_el = $(".chart_column").find("a:first").attr("id");

    if (col_el > 0) {
        table.find("tbody td:nth-child(" + col_el + ")").each(function() {
			columndata.push($(this).text());
        });
    }

    if (row_el > 0) {
        table.find("tbody td:nth-child(" + row_el + ")").each(function() {
			rowdata.push($(this).text());
        });
    }

	var col_total =  0;
	for(var i=0;i<columndata.length;i++)
	{                  
		if(isNaN(columndata[i])) {
	    	continue;
	 	}
		col_total += Number(columndata[i]);
	}

	var chartarea_width = 500;
	
	if(columndata.length > 0 && rowdata.length > 0) {
		switch(charttype) {
			case "bar":
				draw_bar_chart(rowdata, columndata);
			break;
			case "pie":
				draw_pie_chart(rowdata, columndata, col_total);
			break;
			case "column":
				draw_column_chart(rowdata, columndata, col_total);
			break;
			case "doughnut":
				draw_doughnut_chart(rowdata, columndata, col_total);
			break;
			case "area":
				draw_area_chart(rowdata, columndata, col_total);
			break;
			case "line":
				draw_line_chart(rowdata, columndata, col_total);
			break;
		}
	}
}

function getRandomColor() {
    var letters = '0123456789ABCDEF';
    var color = '#';
    for (var i = 0; i < 6; i++) {
        color += letters[Math.floor(Math.random() * 16)];
    }
	if(color == "#FFFFFF")
		getRandomColor();
	else
    	return color;
}

function draw_bar_chart(rowdata, columndata) {
	var bardata = "";
	var counter = 50;
	var chart_height = 32 * rowdata.length;
	
	var max_value = Math.max.apply(Math, columndata);
	
	for(i = 0; i < rowdata.length; i++) {
		var bar_width = Math.ceil((parseInt(columndata[i]) / max_value) * 100);
		
		var translate = parseInt(counter) + 15;
		var translate_data = 500 * bar_width / 100;
		var translate_data2 = counter + 15;
		
		bardata += '<g id="'+rowdata[i]+'" role="listitem">'
				+'<rect class="bar" x="10" y="'+counter+'" width="'+bar_width+'%" height="20" role="presentation" stroke="#0099FF" fill="#0099FF"></rect>'
				+'<text class="series" transform="translate(0.1 '+translate+')" role="presentation" style="text-anchor: end">'+rowdata[i]+'</text>'
				+'<text class="data" transform="translate('+translate_data+' '+translate_data2+')" style="text-anchor: end" fill="white">'+columndata[i]+'</text>'
			+'</g>';
			
		counter = counter + 30;
	}
	bardata = '<g id="bars" role="list" aria-label="bar graph">' + bardata + '</g>';
	bardata = bardata + '<line class="af" x1="42.2" y1="'+chart_height+'" x2="42.2" y2="43.4" role="presentation"></line>';
	bardata = '<svg width="500" height="'+chart_height+'" viewBox="0 0 500 '+chart_height+'">' + bardata + '</svg>';
	
	$(".graph_container_graphs").append(bardata);
}

function draw_pie_chart(rowdata, columndata, col_total) {
	
	var pieslice_json = "";
	var slices = [];
	var legend_table = "";
	var pie_color = [];
	
	for(i = 0; i < rowdata.length; i++) {
		var pie_size = ((parseInt(columndata[i]) / col_total).toFixed(2));
		
		var piecolor = "";
		do {
			piecolor = getRandomColor();
		}while(jQuery.inArray(piecolor, pie_color) !== -1)
		
		
		let pie_size_el = {
		  percent: pie_size,
		  color: piecolor
		};
		slices.push(pie_size_el);
		
		legend_table += '<tr><td style="background-color:'+piecolor+'">'+rowdata[i]+ ' - '+ pie_size +'%</td></tr>';
	}
	legend_table = '<table class="tbl_pie_legends">'+legend_table+'</table>';
	
	var svgel = '<svg id="html2excel_pie_chart" viewBox="-1 -1 2 2" style="transform: rotate(-90deg)" width="500" height="500"></svg>';
	var div_pie = '<div class="row"><div class="col-md-9">'+svgel+'</div><div class="col-md-3">'+legend_table+'</div></div>';
	$(".graph_container_graphs").append(div_pie);	
	
	let cumulativePercent = 0;
    
	function getCoordinatesForPercent(percent) {
	  	const x = Math.cos(2 * Math.PI * percent);
	  	const y = Math.sin(2 * Math.PI * percent);
	  	return [x, y];
	}
    
	var svgpath = "";
	
	slices.forEach(slice => {
	  	const [startX, startY] = getCoordinatesForPercent(cumulativePercent);
	  	cumulativePercent = parseFloat(cumulativePercent) + parseFloat(slice.percent);
	  	const [endX, endY] = getCoordinatesForPercent(cumulativePercent);
	  	const largeArcFlag = slice.percent > .5 ? 1 : 0;
	  	const pathData = [
	  	  `M ${startX} ${startY}`,
	  	  `A 1 1 0 ${largeArcFlag} 1 ${endX} ${endY}`,
	  	  `L 0 0`, // Line
	  	].join(' ');
    
	  	const pathEl = document.createElementNS('http://www.w3.org/2000/svg', 'path');
	  	pathEl.setAttribute('d', pathData);
	  	pathEl.setAttribute('fill', slice.color);
	
		$("#html2excel_pie_chart").append(pathEl);
	});
}

function draw_column_chart(rowdata, columndata, col_total) {
	
	var total_bars = columndata.length;
	var bar_width = 40;
	var svg_width = (parseInt(total_bars) * parseInt(bar_width)) + (parseInt(total_bars) * 2);
	
	var max_bar_height = Math.max.apply(null, columndata) == -Infinity ? 0 : Math.max.apply(null, columndata);
	max_bar_height = parseFloat(max_bar_height) + 50;
	
	var labelxval = parseInt(max_bar_height) + 5;
	
	var xaxis_line = '<line x1="'+bar_width+'" x2="'+svg_width+'" y1="'+max_bar_height+'" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	var yaxis_line = '<line x1="'+bar_width+'" x2="'+bar_width+'" y1="10" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	
	var xaxis_label = "";
	var xlabel_start = 70
	for(i = 0; i < rowdata.length; i++) {
		xaxis_label += '<g><text y="'+xlabel_start+'" x="-'+labelxval+'" style="text-anchor: end;" transform="rotate(-90)">'+rowdata[i]+'</text></g>';
		xlabel_start = parseInt(xlabel_start) + 40;
	}
	xaxis_label = '<g class="col_x_axis">'+xaxis_label+'</g>';
	
	var yaxis_grid = "";
	var total_grid_lines = 10;
	var grid_line_gap = max_bar_height / total_grid_lines;
	var y_val = 25;
	for(i = 0; i < total_grid_lines - 1; i++) {
		yaxis_grid += '<line x1="'+bar_width+'" x2="'+svg_width+'" y1="'+y_val+'" y2="'+y_val+'" style="stroke:#E5E8E8;stroke-width:1"></line>';
		y_val = parseInt(y_val) + parseInt(grid_line_gap);
	}
	yaxis_grid = '<g class="col_y_grid_line">'+yaxis_grid+'</g>';
	
	var rect = "";
	var xaxis_bar_text = "";
	var x2 = 42;
	for(i = 0; i < rowdata.length; i++) {
		var y = parseInt(max_bar_height) -  parseInt(columndata[i]);
		rect += '<rect width="40" height="'+columndata[i]+'" x="'+x2+'" y="'+y+'" style="fill:#0099FF;"></rect>';
		
		var newx = parseInt(x2)+parseInt(10);
		var newy = parseInt(y)-parseInt(5);
		xaxis_bar_text += '<text x="'+newx+'" y="'+newy+'">'+columndata[i]+'</text>';
		
		x2 = parseInt(x2) + 42;
	}
	
	var svg = '<svg id="svg" width="'+svg_width+'px" height="'+max_bar_height+'px">'+xaxis_line+yaxis_line+yaxis_grid+xaxis_label+rect+xaxis_bar_text+'</svg>';
	$(".graph_container_graphs").append(svg);
}

function draw_doughnut_chart(rowdata, columndata, col_total) {
	var radius = 15;
	var cx = 21;
	var cy = 21;
	
	var hole = '<g class="donut-hole"><circle cx="'+cx+'" cy="'+cy+'" r="'+radius+'" fill="#fff"></circle></g>';
	var ring = '<g class="donut-ring"><circle cx="'+cx+'" cy="'+cy+'" r="'+radius+'" fill="transparent" stroke="#d2d3d4" stroke-width="3"></circle></g>';
	
	var segment = "";
	var legend_table = "";
	var offset = 25;
	for(i = 0; i < rowdata.length; i++) {
		var pie_size = ((parseInt(columndata[i]) / col_total).toFixed(2)) * 100;
		var pie_size_balance = parseFloat(100) - parseFloat(pie_size);
		var piecolor = getRandomColor();
		
		if(i > 0)
			offset = parseInt(offset) + parseFloat(pie_size);
		
		segment += '<circle data-per="'+pie_size+'" cx="'+cx+'" cy="'+cy+'" r="'+radius+'" fill="transparent" stroke="'+piecolor+'" stroke-width="3" stroke-dasharray="'+pie_size+' '+pie_size_balance+'" stroke-dashoffset="'+offset+'"></circle>';
		
		legend_table += '<tr><td style="background-color:'+piecolor+'">'+rowdata[i]+ ' - '+ pie_size +'%</td></tr>';
	}
	segment = '<g class="donut-segment">'+segment+'</g>';
	legend_table = '<table class="tbl_pie_legends">'+legend_table+'</table>';

	var svg = '<svg width="300px" height="300px" viewBox="0 0 42 42" class="donut">'+hole+ring+segment+'</svg>';
	
	var doughnut_chart = '<div class="row"><div class="col-md-9">'+svg+'</div><div class="col-md-3">'+legend_table	+'</div></div>';
	
	$(".graph_container_graphs").append(doughnut_chart);
}

function draw_line_chart(rowdata, columndata, col_total) {
	
	var total_bars = columndata.length;
	var bar_width = 40;
	var svg_width = (parseInt(total_bars) * parseInt(bar_width)) + (parseInt(total_bars) * 2);
	
	var max_bar_height = Math.max.apply(null, columndata) == -Infinity ? 0 : Math.max.apply(null, columndata);
	max_bar_height = parseFloat(max_bar_height) + 50;
	
	var labelxval = parseInt(max_bar_height) + 5;
	
	var xaxis_line = '<line x1="'+bar_width+'" x2="'+svg_width+'" y1="'+max_bar_height+'" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	var yaxis_line = '<line x1="'+bar_width+'" x2="'+bar_width+'" y1="10" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	
	var xaxis_label = "";
	var points = "";
	var axis_text = "";
	var pointstart = 40;
	var xlabel_start = 70
	for(i = 0; i < rowdata.length; i++) {
		xaxis_label += '<g><text y="'+xlabel_start+'" x="-'+labelxval+'" style="text-anchor: end;" transform="rotate(-90)">'+rowdata[i]+'</text></g>';
		xlabel_start = parseInt(xlabel_start) + 40;
		
		var value = parseInt(max_bar_height) - parseFloat(columndata[i])
		points += pointstart + " " + value + ",";
		
		axis_text += '<text x="'+pointstart+'" y="'+value+'">'+columndata[i]+'</text>';
		
		pointstart = parseInt(pointstart) + parseInt(40);
	}
	xaxis_label = '<g class="col_x_axis">'+xaxis_label+'</g>';
	var polyline = '<polyline fill="none" stroke="#0099FF" stroke-width="2" points="'+points+'" />';
		
	var svg = '<svg id="svg" width="'+svg_width+'px" height="'+max_bar_height+'px">'+xaxis_line+yaxis_line+xaxis_label+polyline+axis_text+'</svg>';
	$(".graph_container_graphs").append(svg);
}

function draw_area_chart(rowdata, columndata, col_total) {
	
	var total_bars = columndata.length;
	var bar_width = 40;
	var svg_width = (parseInt(total_bars) * parseInt(bar_width)) + (parseInt(total_bars) * 2);
	
	var max_bar_height = Math.max.apply(null, columndata) == -Infinity ? 0 : Math.max.apply(null, columndata);
	max_bar_height = parseFloat(max_bar_height) + 50;
	
	var labelxval = parseInt(max_bar_height) + 5;
	
	var xaxis_line = '<line x1="'+bar_width+'" x2="'+svg_width+'" y1="'+max_bar_height+'" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	var yaxis_line = '<line x1="'+bar_width+'" x2="'+bar_width+'" y1="10" y2="'+max_bar_height+'" style="stroke:#E5E8E8;stroke-width:2"></line>';
	
	var xaxis_label = "";
	var points = "";
	var axis_text = "";
	var pointstart = 40;
	var xlabel_start = 70
	for(i = 0; i < rowdata.length; i++) {
		xaxis_label += '<g><text y="'+xlabel_start+'" x="-'+labelxval+'" style="text-anchor: end;" transform="rotate(-90)">'+rowdata[i]+'</text></g>';
		xlabel_start = parseInt(xlabel_start) + 40;
		
		var value = parseInt(max_bar_height) - parseFloat(columndata[i])
		points += pointstart + " " + value + ",";
		
		axis_text += '<text x="'+pointstart+'" y="'+value+'">'+columndata[i]+'</text>';
		
		pointstart = parseInt(pointstart) + parseInt(40);
	}
	xaxis_label = '<g class="col_x_axis">'+xaxis_label+'</g>';
	var polyline = '<polyline fill="#ccc" stroke="#0099FF" stroke-width="2" points="'+points+'" />';
		
	var svg = '<svg id="svg" width="'+svg_width+'px" height="'+max_bar_height+'px">'+xaxis_line+yaxis_line+xaxis_label+polyline+axis_text+'</svg>';
	$(".graph_container_graphs").append(svg);
}

// END
// GRAPHS 2.0


// EXPORT TABLE 2.0
// START

function getIEVersion() {
    var rv = -1; // Return value assumes failure.
    if (navigator.appName == 'Microsoft Internet Explorer') {
        var ua = navigator.userAgent;
        var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
        if (re.exec(ua) != null)
            rv = parseFloat(RegExp.$1);
    }
    return rv;
}

function tableToExcel(table, sheetName, fileName) {
    var ua = window.navigator.userAgent;
    var msie = ua.indexOf("MSIE ");
    if (msie > 0 || !!navigator.userAgent.match(/Trident.*rv\:11\./)) // If Internet Explorer
    {
        return fnExcelReport(table, fileName);
    }

    var uri = 'data:application/vnd.ms-excel;base64,',
        templateData = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml><![endif]--><meta http-equiv="content-type" content="text/plain; charset=UTF-8"/></head><body><table>{table}</table></body></html>',
        base64Conversion = function(s) {
            return window.btoa(unescape(encodeURIComponent(s)))
        },
        formatExcelData = function(s, c) {
            return s.replace(/{(\w+)}/g, function(m, p) {
                return c[p];
            })
        }

    $("tbody > tr[data-level='0']").show();

    if (!table.nodeType)
        table = document.getElementById(table)

    var ctx = {
        worksheet: sheetName || 'Worksheet',
        table: table.innerHTML
    }

    var element = document.createElement('a');
    element.setAttribute('href', 'data:application/vnd.ms-excel;base64,' + base64Conversion(formatExcelData(templateData, ctx)));
    element.setAttribute('download', fileName);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);

    $("tbody > tr[data-level='0']").hide();
}

function fnExcelReport(table, fileName) {
    var tab_text = "<table border='2px'>";
    var textRange;

    if (!table.nodeType)
        table = document.getElementById(table)

    $("tbody > tr[data-level='0']").show();
    tab_text = tab_text + table.innerHTML;

    tab_text = tab_text + "</table>";
    tab_text = tab_text.replace(/<A[^>]*>|<\/A>/g, ""); //remove if u want links in your table
    tab_text = tab_text.replace(/<img[^>]*>/gi, ""); // remove if u want images in your table
    tab_text = tab_text.replace(/<input[^>]*>|<\/input>/gi, ""); // reomves input params

    txtArea1.document.open("txt/html", "replace");
    txtArea1.document.write(tab_text);
    txtArea1.document.close();
    txtArea1.focus();
    sa = txtArea1.document.execCommand("SaveAs", false, fileName);
    $("tbody > tr[data-level='0']").hide();
    return (sa);
}

function selectElementContents(el) {
    var body = document.body,
        range, sel;
    if (document.createRange && window.getSelection) {
        range = document.createRange();
        sel = window.getSelection();
        sel.removeAllRanges();
        try {
            range.selectNodeContents(el);
            sel.addRange(range);
        } catch (e) {
            range.selectNode(el);
            sel.addRange(range);
        }
    } else if (body.createTextRange) {
        range = body.createTextRange();
        range.moveToElementText(el);
        range.select();
    }
    document.execCommand("Copy");
    return false;
}

function downloadAsJson($this) {
    // Set options
    var defaults = {
        ignoreColumns: [],
        onlyColumns: null,
        ignoreHiddenRows: false,
        headings: null,
        allowHTML: false
    };
    opts = $.extend(defaults, $this);

    var notNull = function(value) {
        return value !== undefined && value !== null;
    };

    var ignoredColumn = function(index) {
        if (notNull(opts.onlyColumns)) {
            return $.inArray(index, opts.onlyColumns) === -1;
        }
        return $.inArray(index, opts.ignoreColumns) !== -1;
    };

    var arraysToHash = function(keys, values) {
        var result = {},
            index = 0;
        $.each(values, function(i, value) {
            if (index < keys.length && notNull(value)) {
                result[keys[index]] = value;
                index++;
            }
        });
        return result;
    };

    var cellValues = function(cellIndex, cell) {
        var value, result;
        if (!ignoredColumn(cellIndex)) {
            var override = $(cell).data('override');
            if (opts.allowHTML) {
                value = $.trim($(cell).html());
            } else {
                value = $.trim($(cell).text());
            }
            result = notNull(override) ? override : value;
        }
        return result;
    };

    var rowValues = function(row) {
        var result = [];
        $(row).children('td,th').each(function(cellIndex, cell) {
            if (!ignoredColumn(cellIndex)) {
                var cVal = cellValues(cellIndex, cell);
                cVal = cVal.replace("↑↓", "");
                result.push(cVal);
            }
        });
        return result;
    };

    var getHeadings = function(table) {
        var firstRow = table.find('tr:first').first();
        return notNull(opts.headings) ? opts.headings : rowValues(firstRow);
    };

    var construct = function(table, headings) {
        var i, j, len, len2, txt, $row, $cell,
            tmpArray = [],
            cellIndex = 0,
            result = [];
        table.children('tbody,*').children('tr').each(function(rowIndex, row) {
            if (rowIndex > 0 || notNull(opts.headings)) {
                $row = $(row);
                if ($row.is(':visible') || !opts.ignoreHiddenRows) {
                    if (!tmpArray[rowIndex]) {
                        tmpArray[rowIndex] = [];
                    }
                    cellIndex = 0;
                    $row.children().each(function() {
                        if (!ignoredColumn(cellIndex)) {
                            $cell = $(this);

                            // process rowspans
                            if ($cell.filter('[rowspan]').length) {
                                len = parseInt($cell.attr('rowspan'), 10) - 1;
                                txt = cellValues(cellIndex, $cell, []);
                                for (i = 1; i <= len; i++) {
                                    if (!tmpArray[rowIndex + i]) {
                                        tmpArray[rowIndex + i] = [];
                                    }
                                    tmpArray[rowIndex + i][cellIndex] = txt;
                                }
                            }
                            // process colspans
                            if ($cell.filter('[colspan]').length) {
                                len = parseInt($cell.attr('colspan'), 10) - 1;
                                txt = cellValues(cellIndex, $cell, []);
                                for (i = 1; i <= len; i++) {
                                    // cell has both col and row spans
                                    if ($cell.filter('[rowspan]').length) {
                                        len2 = parseInt($cell.attr('rowspan'), 10);
                                        for (j = 0; j < len2; j++) {
                                            tmpArray[rowIndex + j][cellIndex + i] = txt;
                                        }
                                    } else {
                                        tmpArray[rowIndex][cellIndex + i] = txt;
                                    }
                                }
                            }
                            // skip column if already defined
                            while (tmpArray[rowIndex][cellIndex]) {
                                cellIndex++;
                            }
                            if (!ignoredColumn(cellIndex)) {
                                txt = tmpArray[rowIndex][cellIndex] || cellValues(cellIndex, $cell, []);
                                if (notNull(txt)) {
                                    tmpArray[rowIndex][cellIndex] = txt;
                                }
                            }
                        }
                        cellIndex++;
                    });
                }
            }
        });
        $.each(tmpArray, function(i, row) {
            if (notNull(row)) {
                txt = arraysToHash(headings, row);
                result[result.length] = txt;
            }
        });
        return result;
    };

    // Run
    var headings = getHeadings($this);
    var jsonoutput = JSON.stringify(construct($this, headings));

    var blob = new Blob([jsonoutput], {
        type: 'json'
    });
    if (window.navigator.msSaveOrOpenBlob) {
        window.navigator.msSaveBlob(blob, "table_controller.json");
    } else {
        var elem = window.document.createElement('a');
        elem.href = window.URL.createObjectURL(blob);
        elem.download = "table_controller.json";
        document.body.appendChild(elem);
        elem.click();
        document.body.removeChild(elem);
    }
}

// END
// EXPORT TABLE 2.0



