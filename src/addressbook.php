<?php

/**
 * addressbook.php
 *
 * Copyright (c) 1999-2003 The SquirrelMail Project Team
 * Licensed under the GNU GPL. For full terms see the file COPYING.
 *
 * Manage personal address book.
 *
 * $Id: addressbook.php,v 1.50 2011/10/21 18:18:32 richb Exp $
 */

/* Path for SquirrelMail required files. */
define('SM_PATH','../');

/* Fix IE6 bug (this is required to send filename in Content-Disposition) */
session_cache_limiter('private');

/* SquirrelMail required files. */
require_once(SM_PATH . 'include/validate.php');
require_once(SM_PATH . 'functions/global.php');
require_once(SM_PATH . 'functions/display_messages.php');
require_once(SM_PATH . 'functions/addressbook.php');
require_once(SM_PATH . 'functions/strings.php');
require_once(SM_PATH . 'functions/html.php');

/*
 * Implementation notes by Rich Braun
 *
 * - Might want to have a separate record-view display (in single screen)
 * - Should have format validation on dates, numbers, and quoted strings.
 * - Test/fix hooks to add addresses (from vcards, message headers, etc).
 * - Get this to work standalone without Squirrelmail.
 * - Get this to work as a loadable Squirrelmail plugin.
 * - Shared/global address book support (with r/w permissions).
 * - Support multiple address books; move items across address books.
 * - Need primary category field for compatibility with Palm.
 * - When creating a record, don't fail if uid is already in use.
 * - Fix the logic which is supposed to preserve entered values on errors.
 * - Get RTF syntax to work with at least one viewer other than Word97.
 * - Is there a way to register this as the default Windows contact manager?
 * - Clean up UI with javascript menus a la winterhillbank.com.
 * - Page footer (print date) needs a bit of work.
 * - Cleanup:  get rid of Street2/Street3 fields.
 * - Filtering and search features.
 * - Address list display needs optimization for large tables.
 * - Import from outlook express:  look at name if lastname/firstname blank;
 *   run phone numbers through formatter.
 * - City/state parser fails for two-word city name without state or ZIP;
 *   "Makati City" becomes "Makati Ci, TY".
 * - Can't set mailing address to Other.
 * - Page footers are missing in birthday list (non-full form factors)
 *   and white-pages (non-full form factors after first page).
 *
 * - The address pop-up is not consistent with the address list.
 * - Would like birthday reminders.
 * - Unused fields: Middle name, title, suffix, nickname, gender, department,
 *   street2/street3, assistants name, assistants phone, managers name,
 *   company main phone, office location, account, hobby, photo.
 * - Value of selcat[] should be cleared after forms-modify.
 */

define('DEF_COUNTRY', "United States of America");
define('DEF_CUSTOM_LABELS', "'Custom 1','Custom 2','Custom 3','Custom 4'");
define('DEF_LIST_COLUMNS',"Name,Email,Company,Phone");
define('NULL_YEAR', "2500");

/* lets get the global vars we may need */
sqgetGlobalVar('key',       $key,           SQ_COOKIE);

sqgetGlobalVar('username',  $username,      SQ_SESSION);
sqgetGlobalVar('onetimepad',$onetimepad,    SQ_SESSION);
sqgetGlobalVar('base_uri',  $base_uri,      SQ_SESSION);
sqgetGlobalVar('delimiter', $delimiter,     SQ_SESSION);

/* From the address form */
sqgetGlobalVar('command',   $command,   SQ_POST);
sqgetGlobalVar('addaddr',   $addaddr,   SQ_POST);
sqgetGlobalVar('editaddr',  $editaddr,  SQ_POST);
sqgetGlobalVar('sel',       $sel,       SQ_POST);
sqgetGlobalVar('oldnick',   $oldnick,   SQ_POST);
sqgetGlobalVar('backend',   $backend,   SQ_POST);
sqgetGlobalVar('doedit',    $doedit,    SQ_POST); 
sqgetGlobalVar('selcat',    $selcat,    SQ_POST);
sqgetGlobalVar('prtfmt',    $prtfmt,    SQ_POST);
sqgetGlobalVar('newcat',    $newcat,    SQ_POST);

/* From the URI */
sqgetGlobalVar('start',     $start);
sqgetGlobalVar('list_numrows', $list_numrows);
sqgetGlobalVar('category',  $active_cat);
sqgetGlobalVar('edituid',   $edituid);

$phonelabels = array( 'Work', 'Home', 'Fax', 'Other', 'E-mail',
		      'Main', 'Pager', 'Cell', 'Work2', 'Home2' );

$printformats = array ( '3-column Detail', 'White Pages', 'Table',
		      'Envelopes #10', 'Envelopes #6', '#5371 Bus. Cards',
		      '#5160 Labels', 'Rolodex 2.25x4', 'Fax Coversheet',
		      'DayRunner 8.5x5.5', 'DayRunner 6.7x3.7', 'CSV',
		      'Vcards', 'Birthday List', 'Nokia CSV', 'Motorola CSV' );
$monthnames  = array ("January", "February", "March", "April", "May",
			"June", "July", "August", "September", "October",
			"November", "December");

/*************************
 * Preferences
 *************************/
if (empty($list_numrows))
  $list_numrows = getPref($data_dir, $username, 'addr_list_numrows');
$r = getpref($data_dir, $username, 'addr_custom_labels');
if (empty($r))
  $r = DEF_CUSTOM_LABELS;
$customlabels = explode ("','", "'," . substr($r, 0, strlen($r)-1));
$r = getPref ($data_dir, $username, 'addr_list_columns');
if (!empty($r))
  $list_columns = explode(",",$r);
else
  $list_columns = explode(",",DEF_LIST_COLUMNS);
$hide_other = getPref($data_dir, $username, 'addr_hide_other');
$lastname_first = getPref($data_dir, $username, 'addr_lastname_first');
$highlight_cat =  strtolower(getPref($data_dir, $username, 'addr_highlight_cat'));

/* (Note:  report options are listed below in the reporting section) */
  
/***************************************
 * Functions
 ***************************************/

/* Make an input field (single line text) */
function input_single_line($label, $field, $name, $size, $values, $add) {
    global $color;
    $td_str = '<INPUT NAME="' . $name . '[' . $field . ']" STYLE="FONT-FAMILY:Arial,proportional;FONT-SIZE:smaller" SIZE="' . $size . '" VALUE="';
    if (isset($values[$field])) {
        $td_str .= htmlspecialchars( strip_tags( $values[$field] ) );
    }
    if (empty($label))
      $td_str .= '" TYPE="hidden">' . $add . '';
    else
      $td_str .= '">' . $add . '';
    return html_tag( 'tr' ,
        html_tag( 'td', empty($label) ? '' : $label . ':', 'right', $color[4]) .
        html_tag( 'td', $td_str, 'left', $color[4])
        )
    . "\n";
}

/* Make an input field (multiple line text) */
function input_multi_line($label, $field, $name, $size, $rows, $values, $add) {
    global $color;
    $td_str = '<TEXTAREA STYLE="FONT-FAMILY:Arial,proportional" COLS="' . $size . '" ROWS="' . $rows . '" NAME="' . $name . '[' . $field . ']">';
    if (isset($values[$field])) {
        $td_str .= htmlspecialchars( strip_tags( $values[$field] ) );
    }
    $td_str .= '</TEXTAREA>' . $add . '';
    return html_tag( 'tr' ,
        html_tag( 'td', $label . ':', 'right', $color[4]) .
        html_tag( 'td', $td_str, 'left', $color[4])
        )
    . "\n";
}

/* Display pick list of phone labels */
function phonepicklist($num, $name, $selected)
{
  global $phonelabels;

  $td_str = '<select name="'.$name.'[phone'.$num.'type]" size=1>\n';
  for ($j = 0; $j < 10; $j++)
  {
    if ($j == $selected)
      $opt = "option selected";
    else
      $opt = "option";

    $td_str .= "<$opt value=$j>".$phonelabels[$j]."</option>\n";
  }
  return $td_str . "</select>";
}

/* Phone # input field (with label drop-down pick list) */
function phone_inp_field($label, $num, $name, $size, $values, $add)
{
  global $color;

  $checked = ($num == $values['phonesel']) ? ' checked' : '';
  $td_str1 = "<class='edit' align=right>\n" . 
             phonepicklist($num, $name, $values["phone".$num."type"]);

  $td_str2 = '<INPUT NAME="' . $name . '[phone' . $num . ']" SIZE="' . $size . '" VALUE="';
    if (isset($values["phone".$num])) {
        $td_str2 .= htmlspecialchars( strip_tags( $values["phone".$num] ) );
    }
    $td_str2 .= '">' . $add . '';
    $td_str2 .= "<input type='radio' name='".$name."[phonesel]' value=".$num.$checked."></input>";
    return html_tag( 'tr' ,
        html_tag( 'td', $td_str1, 'right', $color[4]) .
        html_tag( 'td', $td_str2, 'left', $color[4])
        )
    . "\n";
}

/* Street-address input field (with radio button at right) */
function street_inp_field($label, $field, $name, $size, $values, $add)
{
    global $color;

    if (empty($label))
      return '<INPUT TYPE="hidden" NAME="'.$name.'['.$field.']" VALUE="'.
		htmlspecialchars( strip_tags( $values[$field])).'">';

    $checked = ($values['addresssel'] == substr($label,0,1)) ? ' checked' : '';
    $td_str = '<TEXTAREA STYLE="FONT-FAMILY:Arial,proportional;FONT-SIZE:smaller" COLS="' . $size . '" ROWS="2" NAME="' . $name . '[' . $field . ']">';

    if (isset($values[$field])) {
        $td_str .= htmlspecialchars( strip_tags( $values[$field] ) );
    }
    $td_str .= '</TEXTAREA>' . $add . '';

    $td_str .= "<input type='radio' name='".$name."[addresssel]' value=" .
		substr($label,0,1) . $checked . "></input>";
    $map_url = 'http://www.mapquest.com/maps/map.adp?country=US&address=' .
		rawurlencode($values[strtolower($label).'street']) . '&city=' .
		rawurlencode($values[strtolower($label).'city']) . '&state=' .
		rawurlencode($values[strtolower($label).'state']) . '&zip=' .
		rawurlencode($values[strtolower($label).'zip']) .
		'&homesubmit=Get+Map';
    return html_tag( 'tr' ,
        html_tag( 'td', "<a href=$map_url target=_new>$label Address</a>" .
		  ':', 'right', $color[4]) .
        html_tag( 'td', $td_str, 'left', $color[4])
        )
    . "\n";
}

/* Email-address input field (with radio button at right) */
function email_inp_field($label, $num, $name, $size, $values, $add)
{
    global $color;

    $checked = ($num == $values['emailsel']) ? ' checked' : '';
    $td_str = '<INPUT NAME="' . $name . '[email' . $num . ']" SIZE="' . $size . '" VALUE="';
    if (isset($values['email'.$num])) {
        $td_str .= htmlspecialchars( strip_tags( $values['email'.$num] ) );
    }
    $td_str .= '">' . $add . '';

    $td_str .= "<input type='radio' name='".$name."[emailsel]' value=".$num.$checked."></input>";
    return html_tag( 'tr' ,
        html_tag( 'td', $label . ':', 'right', $color[4]) .
        html_tag( 'td', $td_str, 'left', $color[4])
        )
    . "\n";
}

/* Categories pick list */
function categories_inp_field($label, $num, $name, $size, $which_active, $all, $abook)
{
  global $color;

  $r = $abook->category_get($abook->localbackend);

  if ($all == 1) {
    $td_str = "<select name='selcat[]' size=1>\n";
    $td_str .= "<option value=''>--Categories--</option>\n";
  }
  else {
    $td_str = "<select name='selcat[]' multiple size=6>\n";
  }

  sort ($r);
  foreach( $r as $j => $cat ) {
    $opt = strstr ($which_active, $cat) ? "option selected" : "option";
    $td_str .= "<$opt value='$cat'>" . $cat . "</option>\n";
  }

  if ($all == 1) {
    $opt = strstr ($which_active, 'Unfiled') ? "option selected" : "option";
    $td_str .= "<$opt value='Unfiled'>Unfiled</option>\n";
    $td_str .= "<option value='#(EdiT)#'>Edit...</option>\n";
    return $td_str . '</select>';
  }
  else {
    return html_tag( 'tr', 
        html_tag( 'td', $label . ':', 'right', $color[4]) .
	html_tag( 'td', $td_str . '</select>', $color[4])
    )
    . "\n";
  }
}

/* Print-forms pick list */
function print_picklist ($printformats)
{
    $td_str = "<select name=prtfmt size=1 onChange=\"javascript:formHandler(addr)\">\n";
    $td_str .= "<option value=''>--Print--</option>\n";

    foreach( $printformats as $format )
	$td_str .= "<option value='print $format'>" . $format . "</option>\n";

    return $td_str . '</select>';
}

function getSmallStringCell($string, $align, $cols) {
    return html_tag('td',
                    '<small>' . $string . '&nbsp; </small>',
                    $align, '', sprintf('colspan=%d', $cols));
}


/* Display revision history */
function show_history ($abook, $record) {
    global $color;

    echo '<TABLE border=0 cellpadding=0 cellspacing=0>' . "\n";
    echo html_tag('tr',
		html_tag( 'td', "Revision History", 'center', $color[9],
			  'colspan=4'), '') . "\n";
    echo html_tag('tr',
		html_tag( 'td', "Modified", 'center', $color[5]) .
		html_tag( 'td', "&nbsp Field", '', $color[5]) .
		html_tag( 'td', "&nbsp To", '', $color[5]) .
		html_tag( 'td', "&nbsp From", '', $color[5]),
	 '') . "\n";
    echo html_tag('tr',
		html_tag( 'td', '', '', $color[3], 'colspan=4 height=1'),
         '');
    echo html_tag ('tr',
		html_tag( 'td', date_fmt_sql($record['created'],'d-M-y H:i'), 'center', $color[4]) .
		html_tag( 'td', '', '', $color[4]) .
		html_tag( 'td', "&nbsp <i>Created</i>", '', $color[4]) .
		html_tag( 'td', '', '', $color[4]),
	 '') . "\n";
    $line = 1;
    if ($record['created'] != $record['modified']) {
	echo html_tag('tr',
		html_tag( 'td', '', '', $color[3], 'colspan=4 height=1'),
	     '');
	echo html_tag ('tr',
		html_tag( 'td', date_fmt_sql($record['modified'],'d-M-y H:i'), 'center', $color[0]) .
		html_tag( 'td', '', '', $color[0]) .
		html_tag( 'td', "&nbsp <i>Modified</i>", '', $color[0]) .
		html_tag( 'td', '', '', $color[0]),
	     '') . "\n";
	$line++;
    }

    $history = $abook->get_history ($record['uid'], $abook->localbackend);
    foreach( $history as $entry ) {
	/* Display thin line between rows */
	echo html_tag('tr',
		html_tag( 'td', '', '', $color[3], 'colspan=4 height=1'),
	     '');

	/* Display an entry */
	$td_color = (++$line % 2) ? $color[4] : $color[0];
	echo html_tag('tr',
		html_tag( 'td', date_fmt_sql($entry['moddate'],'d-M-y H:i'), 'center', $td_color) .
		html_tag( 'td', "&nbsp {$entry['field']}", '', $td_color) .
		html_tag( 'td', "&nbsp {$entry['newval']}", '', $td_color) .
		html_tag( 'td', "&nbsp {$entry['oldval']} &nbsp", '', $td_color),
	     '') . "\n";
    }
    echo html_tag('tr',
		html_tag( 'td', '', '', $color[3], 'colspan=4 height=1'),
         '');
    echo '</TABLE>';	
}

/* Output form to add and modify address data */
function address_form($abook, $name, $submittext, $values = array()) {
    global $color, $squirrelmail_language, $customlabels, $hide_other;
    
    if ($squirrelmail_language == 'ja_JP') {
	echo html_tag( 'table',
                       input_single_line(_("Nickname"),     'uid', $name, 15, $values, '') .
                       input_single_line(_("Last name"),    'LastName', $name, 45, $values, '') .
                       input_single_line(_("First name"),  'FirstName', $name, 45, $values, '') .
                       input_single_line(_("E-mail address"),   'Email1', $name, 45, $values, '') .
                       input_single_line(_("Additional info"), 'Notes', $name, 45, $values, '') .
                       html_tag( 'tr',
                           html_tag( 'td',
                                       '<INPUT TYPE=submit NAME="' . $name . '[SUBMIT]" VALUE="' .
                                       $submittext . '">',
                                   'center', $color[4], 'colspan="2"')
                       )
	    , 'center', '', 'border="0" cellpadding="1" width="90%"') ."\n";
    }
    else {
	$tr_custom = '';
	for ($i = 1; $i <= 4; $i++) {
	    $tr_custom .= input_single_line(_($customlabels[$i]),
		    'user'.$i, $name, 45, $values, '');
	}
	$values['birthday'] = date_fmt($values['birthday'], 'dd-mmm-yy');
	$values['anniversary'] = date_fmt($values['anniversary'], 'dd-mmm-yy');

	echo html_tag( 'table',
           input_single_line(_(""),       'uid', $name, 15, $values, '') .
           input_single_line(_("First Name"),  'firstname', $name, 45, $values, '') .
           input_single_line(_("Last name"),   'lastname', $name, 45, $values, '') .
           input_single_line(_("Title"),       'jobtitle', $name, 45, $values, '') .
           input_single_line(_("Company"),     'company', $name, 45, $values, '') .
	   phone_inp_field(_("Phone1"), '1', $name, 45, $values, '') .
	   phone_inp_field(_("Phone2"), '2', $name, 45, $values, '') .
	   phone_inp_field(_("Phone3"), '3', $name, 45, $values, '') .
	   phone_inp_field(_("Phone4"), '4', $name, 45, $values, '') .
	   phone_inp_field(_("Phone5"), '5', $name, 45, $values, '') .
           email_inp_field(_("E-mail address"),  '1', $name, 45, $values, '') .
           email_inp_field(_("E-mail 2"),  '2', $name, 45, $values, '') .
           email_inp_field(_("E-mail 3"),  '3', $name, 45, $values, '') .
           street_inp_field(_("Home"), 'homeaddress', $name, 38, $values, '') .
           input_single_line(_("Country"),     'homecountry', $name, 19, $values, '') .
           street_inp_field(_("Business"),     'businessaddress', $name, 38, $values, '') .
           input_single_line(_("Country"),     'businesscountry', $name, 19, $values, '') .
           street_inp_field($hide_other ? '' : _("Other"),        'otheraddress', $name, 38, $values, '') .
           input_single_line($hide_other ? '' : _("Country"),     'othercountry', $name, 19, $values, '') .
           input_single_line(_("Sig. Other"),  'sigother', $name, 45, $values, '') .
           input_single_line(_("Children"),    'children', $name, 45, $values, '') .
           input_single_line(_("Anniversary"), 'anniversary', $name, 7, $values, '') .
           input_single_line(_("Birth date"),  'birthday', $name, 7, $values, '') .
           input_single_line(_("Web Page"),    'webpage', $name, 45, $values, '') .
	   $tr_custom .
	   categories_inp_field(_("Categories"),  'categories', $name, 45, $values['categories'], 0, $abook) .
           input_multi_line(_("Notes"), 'notes', $name, 41, 5, $values, '') .
                       html_tag( 'tr',
                           html_tag( 'td',
                                       '<INPUT TYPE=submit NAME="' . $name . '[SUBMIT]" VALUE="' .
                                       $submittext . '">'. "\n " .
                                       '<INPUT TYPE=submit NAME="command" VALUE="Cancel"' . '">',
                                   'center', $color[4], 'colspan="2"')
                       )
	, 'center', '', 'border="0" cellpadding="1" width="90%"') ."\n";

	if (strstr($submittext, 'Update'))
	    echo "<BR><P>\n" . show_history ($abook, $values);
	else {
	    echo '<TABLE align="center" border=0 bgcolor=FFFFFF cellpadding=0 cellspacing=0 >';
	    echo html_tag( 'tr',
		   html_tag( 'td', 'Import from file<hr>', 'center', $color[0],
			     'HEIGHT="50"')
			 ) .
		 html_tag( 'tr',
		   html_tag( 'td', '<INPUT NAME="userfile" size="35" TYPE="file">',
			     'center', $color[0])
			 ) .
		 html_tag( 'tr',
		   html_tag( 'td', '<INPUT TYPE="submit" VALUE="Upload File" NAME="command">',
			     'center', $color[0])
			 ) .
		 html_tag( 'tr',
		   html_tag( 'td', '<INPUT TYPE="hidden" NAME="directory" VALUE="<?php echo $directory;?>">' .
			    	   '<INPUT TYPE="hidden" NAME="Itemid" VALUE="<?php echo $Itemid;?>">',
			     'center', $color[0])
			 );
	    echo '</TABLE>';
	}
    }
}

/*
 * RTF header
 */
function rtf_header ($footer)
{
    if ($footer)
	$footer_text = '{\footer \pard\plain 
\s16\widctlpar\tqc\tx5040\tqr\tx10440\adjustright \f1\fs20\cgrid {\field{\*\fldinst {\cgrid0  AUTHOR }}{\fldrslt {\lang1024\cgrid0 $username}}}{\cgrid0 \tab Page }{\field{\*\fldinst {\cgrid0  PAGE }}{\fldrslt {\lang1024\cgrid0 1}}}{\cgrid0 \tab }
{\field{\*\fldinst {\cgrid0  DATE }}{\fldrslt {\lang1024\cgrid0 05/14/03}}}{
\par }}';
    else
	$footer_text = '';

    return '{\rtf1\ansi\ansicpg1252\uc1 \deff1\deflang1033\deflangfe1033{\fonttbl
{\f0\froman\fcharset0\fprq2
{\*\panose 02020603050405020304}Times New Roman;}
}
{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;
\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}
{\stylesheet
{\widctlpar\adjustright \f1\fs20\cgrid \snext0 Normal;}
{\*\cs10 \additive Default Paragraph Font;}}\margl864\margr864\margt1008\margb1008 \widowctrl\ftnbj\aenddoc\hyphcaps0\formshade\viewkind1\viewscale140\viewzk2\pgbrdrhead\pgbrdrfoot \fet0\sectd \linex0\cols3\colsx360\endnhere\sectdefaultcl 
'.$footer_text.'{\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}
{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8
\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}\pard\plain \widctlpar\adjustright \f1\fs20\cgrid 
';
}

function rtf_header_labels ($footer)
{
    if ($footer)
	$footer_text = '';
    return '{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f16\froman\fcharset238\fprq2 Times New Roman CE;}{\f17\froman\fcharset204\fprq2 Times New Roman Cyr;}
{\f19\froman\fcharset161\fprq2 Times New Roman Greek;}{\f20\froman\fcharset162\fprq2 Times New Roman Tur;}{\f21\froman\fcharset186\fprq2 Times New Roman Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;
\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;
\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{\widctlpar\adjustright \fs20\cgrid \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}}\margl272\margr272\margt720\margb0 
\widowctrl\ftnbj\aenddoc\formshade\viewkind1\viewscale144\viewzk2\pgbrdrhead\pgbrdrfoot \fet0\sectd \binfsxn4\binsxn4\sbknone\linex0\endnhere\sectdefaultcl {\*\pnseclvl1\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2
\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6
\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang
{\pntxtb (}{\pntxta )}}';

}

function rtf_header_table ($footer)
{
    if ($footer)
	$footer_text = '{\footer \pard\plain 
\s16\widctlpar\tqc\tx5040\tqr\tx10440\adjustright \f1\fs20\cgrid {\field{\*\fldinst {\cgrid0  AUTHOR }}{\fldrslt {\lang1024\cgrid0 $username}}}{\cgrid0 \tab Page }{\field{\*\fldinst {\cgrid0  PAGE }}{\fldrslt {\lang1024\cgrid0 1}}}{\cgrid0 \tab }
{\field{\*\fldinst {\cgrid0  DATE }}{\fldrslt {\lang1024\cgrid0 05/14/03}}}{
\par }}';
    else
	$footer_text = '';

    return '{\rtf1\ansi\ansicpg1252\uc1 \deff0\deflang1033\deflangfe1033{\fonttbl{\f0\froman\fcharset0\fprq2{\*\panose 02020603050405020304}Times New Roman;}{\f28\froman\fcharset238\fprq2 Times New Roman CE;}{\f29\froman\fcharset204\fprq2 Times New Roman Cyr;}
{\f31\froman\fcharset161\fprq2 Times New Roman Greek;}{\f32\froman\fcharset162\fprq2 Times New Roman Tur;}{\f33\froman\fcharset177\fprq2 Times New Roman (Hebrew);}{\f34\froman\fcharset178\fprq2 Times New Roman (Arabic);}
{\f35\froman\fcharset186\fprq2 Times New Roman Baltic;}}{\colortbl;\red0\green0\blue0;\red0\green0\blue255;\red0\green255\blue255;\red0\green255\blue0;\red255\green0\blue255;\red255\green0\blue0;\red255\green255\blue0;\red255\green255\blue255;
\red0\green0\blue128;\red0\green128\blue128;\red0\green128\blue0;\red128\green0\blue128;\red128\green0\blue0;\red128\green128\blue0;\red128\green128\blue128;\red192\green192\blue192;}{\stylesheet{
\ql \li0\ri0\widctlpar\aspalpha\aspnum\faauto\adjustright\rin0\lin0\itap0 \fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 \snext0 Normal;}{\*\cs10 \additive Default Paragraph Font;}}{\info{\title Edward Bryan}{\author swells}{\operator swells}
{\creatim\yr2003\mo5\dy29\hr14\min41}{\revtim\yr2003\mo5\dy29\hr14\min42}{\version1}{\edmins1}{\nofpages1}{\nofwords0}{\nofchars0}{\*\company Decemberfarm}{\nofcharsws0}{\vern8269}}\margl272\margr272\margt720\margb0 
\widowctrl\ftnbj\aenddoc\noxlattoyen\expshrtn\noultrlspc\dntblnsbdb\nospaceforul\formshade\horzdoc\dgmargin\dghspace180\dgvspace180\dghorigin272\dgvorigin720\dghshow1\dgvshow1
\jexpand\viewkind1\viewscale129\viewzk2\pgbrdrhead\pgbrdrfoot\splytwnine\ftnlytwnine\htmautsp\nolnhtadjtbl\useltbaln\alntblind\lytcalctblwd\lyttblrtgr\lnbrkrule \fet0\sectd \binfsxn4\binsxn4\sbknone\linex0\endnhere\sectdefaultcl {\*\pnseclvl1
\pnucrm\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl2\pnucltr\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl3\pndec\pnstart1\pnindent720\pnhang{\pntxta .}}{\*\pnseclvl4\pnlcltr\pnstart1\pnindent720\pnhang{\pntxta )}}{\*\pnseclvl5
\pndec\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl6\pnlcltr\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl7\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}{\*\pnseclvl8\pnlcltr\pnstart1\pnindent720\pnhang
{\pntxtb (}{\pntxta )}}{\*\pnseclvl9\pnlcrm\pnstart1\pnindent720\pnhang{\pntxtb (}{\pntxta )}}';
}

/*
 * Output page form-factor settings
 */
function rtf_page_format ($form_factor, $prt_opt) {
    $td_out = '';
    if ($form_factor == 'pers') {
	$td_out .= '\paperw5400\paperh9720\margl432\margr432\margt432\margb432\gutter360';
        $td_out .= '{\fonttbl{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$prt_opt[form_pers][fontname].';}}';
	$td_out .= '\pgbrdrhead\pgbrdrfoot\pgbrdrsnap \fet0\sectd \linex0\endnhere\pgbrdrt
\brdrtnthsg\brdrw60\brsp20 \pgbrdrl\brdrtnthsg\brdrw60\brsp80 \pgbrdrb\brdrthtnsg\brdrw60\brsp20 \pgbrdrr\brdrthtnsg\brdrw60\brsp80 \sectdefaultcl\cols1\fs'.(($prt_opt[form_pers][fontsize]+1)*2);
    }
    else
    if ($form_factor == 'half') {
	$td_out .= '\paperw7920\paperh12240\margl432\margr432\margt432\margb432\gutter504';
        $td_out .= '{\fonttbl{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$prt_opt[form_half][fontname].';}}';
	$td_out .= '\pgbrdrhead\pgbrdrfoot\pgbrdrsnap \fet0\sectd \linex0\cols2\colsx288\linebetcol\endnhere\pgbrdrt
\brdrtnthsg\brdrw60\brsp20 \pgbrdrl\brdrtnthsg\brdrw60\brsp80 \pgbrdrb\brdrthtnsg\brdrw60\brsp20 \pgbrdrr\brdrthtnsg\brdrw60\brsp80 \sectdefaultcl\fs'.(($prt_opt[form_half][fontsize]+1)*2);
    }
    if ($prt_opt['general']['duplex'])
	$td_out .= '\margmirror';
    return $td_out;
}

/*
 * Quoted-printable encode (like imap_8bit) for 7bit ASCII
 */
function need_qp_enc ($str) {
    if (!strstr ($str, "\n") && strlen ($str) < 76)
	return '';
    return ';ENCODING=QUOTED-PRINTABLE';
}

function qp_enc ($str) {
    if (!strstr ($str, "\n") && strlen ($str) < 76)
	return $str;
    $ret = str_replace("\r\n", "=0D=0A", $str);
    $ret = str_replace("\n", "=0D=0A", $ret);
    return wordwrap($ret, 75, "=\r\n", 1);
}

/*
 * parse_city
 *   Given a multi-line street address, split out the street, city, state,
 *   ZIP fields.  Set the country name to default.
 */

function parse_address ($addr, $addrtype, &$rec)
{
    $pat_zip = '/^([-a-zA-Z ]+),?\s?(\w{2})\s+(\d{5}(-\d{4})?)*$/';
    $pat_nozip = '/^([-a-zA-Z ]+),?\s?(\w{2})$/';
    $addrlines = explode ("\n", trim($addr));
    $cityline = array_pop($addrlines);
    
    if (strlen(trim($addr)) && empty( $rec[$addrtype.'country'] ))
	$rec[$addrtype.'country'] = DEF_COUNTRY;

    if (preg_match( $pat_zip, $cityline, $fields)) {
	$rec[$addrtype.'city']  = $fields[1];
	$rec[$addrtype.'state'] = strtoupper($fields[2]);
	$rec[$addrtype.'zip']   = $fields[3];
	$rec[$addrtype.'street'] = trim(implode ("\n", $addrlines));
	return true;
    }
    else if (preg_match( $pat_nozip, $cityline, $fields)) {
	$rec[$addrtype.'city']  = $fields[1];
	$rec[$addrtype.'state'] = strtoupper($fields[2]);
	$rec[$addrtype.'zip']   = '';
	$rec[$addrtype.'street'] = trim(implode ("\n", $addrlines));
	return true;
    }
    else {
	$rec[$addrtype.'street'] = trim($addr);
	$rec[$addrtype.'city'] = '';
	$rec[$addrtype.'state'] = '';
	$rec[$addrtype.'zip'] = '';
    }
    return empty($rec[$addrtype.'street']) ? true : false;
}

function parse_phone ($number, $index, &$rec)
{
    $pat = '/^(\+?1\-?\s*)?[(]?(\d{3})[)]?-?\s*(\d{3})-?\s*(\d{4})$/';

    if (preg_match( $pat, trim($number), $fields)) {
	$rec['phone'.$index] = "($fields[2]) $fields[3]-$fields[4]";
	return true;
    }
    return empty($number) ? true : false;
}

function parse_date ($dateval, $field, &$rec)
{
    global $monthnames;

    $pat1 = '/^(\d{4}-\d{2}-\d{2})$/';
    $pat2 = '/^(\d\d?)[-\/](\d\d?)[-\/](\d{2})$/';
    $pat3 = '/^(\d\d?)[-\/](\d\d?)/';
    $pat4 = '/^(\d\d?)[- ](\w+)[- ](\d\d*)/';
    $pat5 = '/^(\d\d?)[- ](\w+)/';

    if (!strlen(trim($dateval))) {
	$rec[$field] = '0000-00-00';
	return true;
    }

    if (preg_match( $pat1, trim($dateval), $fields)) {
	$rec[$field] = $fields[1];
	return true;
    }
    else if (preg_match( $pat2, trim($dateval), $fields)) {
	$month = $fields[1];
	$mday = $fields[2];
	$year = $fields[3];	
    }
    else if (preg_match( $pat3, trim($dateval), $fields)) {
	$month = $fields[1];
	$mday = $fields[2];
	$year = NULL_YEAR;
    }
    else if (preg_match( $pat4, trim($dateval), $fields)) {
	$mday = $fields[1];
	foreach ($monthnames as $month => $name)
	    if (!strcasecmp($fields[2],
		substr($name,0,strlen($fields[2])))) break;
	$month = $month + 1;
	$year = $fields[3];
    }
    else if (preg_match( $pat5, trim($dateval), $fields)) {
	$mday = $fields[1];
	foreach ($monthnames as $month => $name)
	    if (!strcasecmp($fields[2],
		substr($name,0,strlen($fields[2])))) break;
	$month = $month + 1;
	$year = NULL_YEAR;
    }
    if (!empty($year) && $mday >= 1 && $mday <= 31 && $year < 10000) {
	if ($year < 20)
	    $year += 2000;
	else if ($year < 100)
	   $year += 1900;
	$rec[$field] = sprintf ('%4d-%02d-%02d', $year, $month, $mday);
	return true;
    }
    $rec[$field] = '0000-00-00';
    return false;
}

function country_is_us ($country) {
    $country = trim($country);
    if (empty($country) ||
	!strcasecmp ($country, 'United States of America') ||
	!strcasecmp ($country, 'United States') ||
	!strcasecmp ($country, 'US') ||
	!strcasecmp ($country, 'USA'))
	return true;
    else
	return false;
}

function date_fmt ($str, $fmt) {
    global $monthnames;
    $year = substr($str,0,4);
    if ($year == 0)
	return '';
    if ($fmt == 'mm-dd-yy') {
	$ret = ltrim(substr($str,5,5), '0');
	if ($year != NULL_YEAR)
	    $ret .= '-' . substr($year, 2, 2);
    }
    else if ($fmt == 'dd-mmm-yy') {
	$ret = ltrim(substr($str,8,2),'0') . '-' . substr($monthnames[substr($str,5,2)-1], 0, 3);
	if ($year != NULL_YEAR)
	    $ret .= '-' . substr($year, 2, 2);
    }	
    else {
	$ret = '!InvDt!';
    }
    return $ret;

}

function date_fmt_sql($sql_timestamp, $fmt)
{
   $month  = substr($sql_timestamp,5,2);
   $day    = substr($sql_timestamp,8,2);
   $year   = substr($sql_timestamp,0,4);

   $hour   = substr($sql_timestamp,11,2);
   $min    = substr($sql_timestamp,14,2);
   $sec    = substr($sql_timestamp,17,2);

   $epoch  = mktime($hour,$min,$sec,$month,$day,$year);
   return date ($fmt, $epoch);
}

function get_sort($dir, $user) {
    $sort = ''; $comma = '';
    for ($i = 1; $i <= 5; $i++) {
	$p1 = getPref ($dir, $user, 'addr_sort_'.$i.'_field');
	if (strlen(str_replace(' ', '', $p1))) {
	    $sort .= $comma . $p1 . " " .
			getPref($dir, $user, 'addr_sort_'.$i.'_dir');
	    $comma = ',';
	}
    }
    if (empty($sort))
	$sort = 'lastname ASC,firstname ASC';
    return $sort;
}

/************************************
 * Java functions
 *  Not yet debugged, has no effect
 ************************************/

$javafunctions = '
<script language="JavaScript" type="text/javascript">
<!-- Begin
function check_all(f) {
 len=f.length;
 for (i=0;i<len;i++) {
  if (f.elements[i].type=="checkbox") {
    f.elements[i].checked=true
  }
 }
}
// End -->

<!-- Begin
function formHandler(form){
var SEL = document.form.command.options[document.form.command.selectedIndex].value;
window.location.href = SEL;
}
// End -->

function validateNumber(field, msg, min, max) {
	if (!min) { min = 0 }
	if (!max) { max = 255 }

	if ( (parseInt(field.value) != field.value) ||
             field.value.length < min ||
             field.value.length > max) {
		alert(msg);
		field.focus();
		field.select();
		return false;
	}

	return true;
}
</script>
';


/************************************
 * Beginning of program
 ************************************/

/* Open addressbook, with error messages on but without LDAP (the *
 * second "true"). Don't need LDAP here anyway                    */
$abook = addressbook_init(true, true);
if($abook->localbackend == 0) {
    plain_error_message(
        _("No personal address book is defined. Contact administrator."),
        $color);
    exit();
}

$defdata   = array();
$formerror = '';
$abortform = false;
$showform  = 'addrlist';
$defselected  = array();
$form_url = 'addressbook.php';
$page_header_sent = false;

if (isset($edituid) && !isset($command)) {
    list($ebackend, $uid) = explode(':', $edituid);
    $defdata = $abook->lookup($uid, $ebackend);
    $showform = 'addrform';
}

/**********************************
 * Execute user command
 **********************************/
if(sqgetGlobalVar('REQUEST_METHOD', $req_method, SQ_SERVER) && $req_method == 'POST') {

    /**************************************************
     * Add new address                                *
     **************************************************/
    if (!empty($addaddr['uid']) && $command != 'Cancel') {
        foreach( $addaddr as $k => $adr ) {
            $addaddr[$k] = strip_tags( $adr );
        }

	/* Collect forms data */
	if (sizeof($selcat) == 0)
	    $addaddr['categories'] = '';
	else
	    $addaddr['categories'] = implode(',', $selcat);

	/* Parse the city/state/ZIP */

	if (!parse_address ($addaddr['homeaddress'], 'home', $addaddr))
	    $formerror = _("Warning, did not recognize city/state/ZIP in home address");
	if (!parse_address ($addaddr['businessaddress'], 'business', $addaddr))
	    $formerror = _("Warning, did not recognize city/state/ZIP in business address");
	if (!parse_address ($addaddr['otheraddress'], 'other', $addaddr))
	    $formerror = _("Warning, did not recognize city/state/ZIP in other address");

	if (!empty($addaddr[homeaddress]) && empty($addaddr[businessaddress]))
	    $addaddr[addresssel] = 'H';

	/* If no selected phone field is blank, look for nonblank one */
	if (empty($addaddr['phone'.$addaddr['phonesel']])) {
	   if (!empty($addaddr['phone2']))
		$addaddr['phonesel'] = 2;  // 1st preference: home phone
	   else if (!empty($addaddr['phone1']))
		$addaddr['phonesel'] = 1;  // 2nd preference: work phone
	   else if (!empty($addaddr['phone4']))
		$addaddr['phonesel'] = 4;  // 3rd preference: cell phone
	}
	for ($i = 1; $i <=5; $i++)
	    parse_phone ($addaddr['phone'.$i], $i, $addaddr);
	parse_date ($addaddr['anniversary'], 'anniversary', $addaddr);
	parse_date ($addaddr['birthday'], 'birthday', $addaddr);

        $r = $abook->add($addaddr, $abook->localbackend);

        /* Handle error messages */
        if (!$r) {
            /* Remove backend name from error string */
            $errstr = $abook->error;
            $errstr = ereg_replace('^\[.*\] *', '', $errstr);

            $formerror = $errstr;
            $showform = 'addrform';
            $defdata = $addaddr;
        }
	else {
	    $showform = 'addrform';
	    $command = 'Create';
	}
    }


if (isset($command)) {
    list( $action, $params ) = explode( ":", $command, 2 );

    switch ($action) {
    case 'EditUID':
    /**************************************************
     * Edit button clicked                            *
     **************************************************/

        list($ebackend, $uid) = explode(':', $params);
	$defdata = $abook->lookup($uid, $ebackend);
	$showform = 'addrform';
	break;

    case 'Create':
    /**************************************************
     * Create address (using form)                    *
     **************************************************/

	$showform = 'addrform';

	/* Set up default values for a new record */
	$defdata['uid'] = "_" . GenerateRandomString( 9, "ABCDEF", 4);
	$defdata['phone2type'] = "1";
	$defdata['phone3type'] = "2";
	$defdata['phone4type'] = "7";
	$defdata['phone5type'] = "3";
	$defdata['phonesel']   = "1";
	$defdata['categories'] = "Unfiled";
	$defdata['emailsel']   = "1";
	$defdata['addresssel'] = "B";
	break;

    case 'Delete':
    /************************************************
     * Delete address(es)                           *
     ************************************************/
	if (sizeof($sel) == 0) {
            $formerror = 'Please select at least one record';
	    break;
	}
	if (!$page_header_sent)
	   displayPageHeader($color, 'None');
	$page_header_sent = true;

	echo '<FORM ACTION="' . $form_url . '" METHOD="POST">' . "\n";
	echo '<CENTER><B>Removing ' . sizeof($sel) . ' record' .
		(sizeof($sel) != 1 ? 's' : '') . '!</B><BR>';
	echo '<INPUT TYPE=submit NAME=command VALUE="' .
                                         _("Click to Confirm") . "\">\n";
	echo '<INPUT TYPE=hidden NAME="sel" VALUE="' .
					 implode(',', $sel) . "\">\n";
	echo '</CENTER></FORM>';
	$abortform = true;
	break;

    case 'Click to Confirm':
	$sel = explode(',', $sel);
        if (sizeof($sel) > 0) {
            $orig_sel = $sel;
            sort($sel);

            /* The selected addresses are identidied by "backend:uid". *
             * Sort the list and process one backend at the time            */
            $prevback  = -1;
            $subsel    = array();
            $delfailed = false;

            for ($i = 0 ; (($i < sizeof($sel)) && !$delfailed) ; $i++) {
                list($sbackend, $uid) = explode(':', $sel[$i]);

                /* When we get to a new backend, process addresses in *
                 * previous one.                                      */
                if ($prevback != $sbackend && $prevback != -1) {

                    $r = $abook->remove($subsel, $prevback);
                    if (!$r) {
                        $formerror = $abook->error;
                        $i = sizeof($sel);
                        $delfailed = true;
                        break;
                    }
                    $subsel   = array();
                }

                /* Queue for processing */
                array_push($subsel, $uid);
                $prevback = $sbackend;
            }

            if (!$delfailed) {
                $r = $abook->remove($subsel, $prevback);
                if (!$r) { /* Handle errors */
                    $formerror = $abook->error;
                    $delfailed = true;
                }
            }

            if ($delfailed) {
                $showform = 'addrlist';
                $defselected  = $orig_sel;
            }
	}
	break;

    case 'Choose':
        /************************************************
         * Apply category filter                        *
         ************************************************/
	if ($selcat[0] == '#(EdiT)#')
	    $showform = 'catmanager';
	else
	    $active_cat = $selcat[0];
	break;

    case 'Insert':
        /************************************************
         * Add selected records to category             *
         ************************************************/
	if(!isset($sel) || (sizeof($sel) == 0)) {
	    $formerror = _("Please select at least one record");
	    break;
        }
	if ($selcat[0] == 'Unfiled') {
	    $formerror = _("Cannot insert to 'Unfiled' category");
	    break;
	}

	list($sbackend, $uid) = explode(':', $sel[0]);

        /* The selected addresses are identified by "backend:uid". *
         * Sort the list and process one backend at a time         */
        $prevback  = -1;
        $subsel    = array();
	$records   = '';
        for ($i = 0 ; $i < sizeof($sel) ; $i++) {
            list($sbackend, $uid) = explode(':', $sel[$i]);

            /* When we get to a new backend, process addresses in *
             * previous one.                                      */
            if ($prevback != $sbackend && $prevback != -1) {
		$r = $abook->category_addrecs($selcat[0], $subsel, $prevback);
		if (!$r) {
		    $formerror = $abook->error;
		    $i = sizeof($sel);
		    $failed = true;
		    break;
		}
                $subsel   = array();
            }

            /* Queue for processing */
            array_push($subsel, $uid);
            $prevback = $sbackend;
        }
	if (!$failed) {
	    $r = $abook->category_addrecs($selcat[0], $subsel, $prevback);
	    if (!$r) {
		$formerror = $abook->error;
		$failed = true;
	    }
	}
//	if ($failed)
//	   $defselected = $orig_sel;
	break;

    case 'Remove':
        /************************************************
         * Remove selected records from category        *
         ************************************************/
	if(!isset($sel) || (sizeof($sel) == 0)) {
	    $formerror = _("Please select at least one record");
            $showform = 'addrlist';
	    break;
        }

	list($sbackend, $uid) = explode(':', $sel[0]);

        /* The selected addresses are identified by "backend:uid". *
         * Sort the list and process one backend at a time         */
        $prevback  = -1;
        $subsel    = array();
	$records   = '';
        for ($i = 0 ; $i < sizeof($sel) ; $i++) {
            list($sbackend, $uid) = explode(':', $sel[$i]);

            /* When we get to a new backend, process addresses in *
             * previous one.                                      */
            if ($prevback != $sbackend && $prevback != -1) {
		$r = $abook->category_rmrecs($selcat[0], $subsel, $prevback);
		if (!$r) {
		    $formerror = $abook->error;
		    $i = sizeof($sel);
		    $failed = true;
		    break;
		}
                $subsel   = array();
            }

            /* Queue for processing */
            array_push($subsel, $uid);
            $prevback = $sbackend;
        }
	if (!$failed) {
	    $r = $abook->category_rmrecs($selcat[0], $subsel, $prevback);
	    if (!$r) {
		$formerror = $abook->error;
		$failed = true;
	    }
	}
//	if ($failed)
//	   $defselected = $orig_sel;
	break;

    case 'Add New Category':
	$showform = 'catmanager';
	if ($newcat == 'Unfiled') {
	   $formerror = 'Invalid category name';
	   break;
	}
	$r = $abook->category_get($abook->localbackend);

	if (!in_array($newcat, $r)) {
	    $r = $abook->category_create($newcat, $abook->localbackend);
	    if (!$r) {
        	/* Remove backend name from error string */
                $errstr = $abook->error;
                $errstr = ereg_replace('^\[.*\] *', '', $errstr);

                $formerror = $errstr;
                $defdata = $addaddr;
            }
        }
	break;

    case 'Delete Categories':
	if (!isset($sel) || (sizeof($sel) == 0)) {
	    $formerror = 'Choose at least one category to delete';
            $orig_sel = $sel;
	}    
	else {
	    $r = $abook->category_drop($sel, $abook->localbackend);
	    if (!$r) {
        	/* Remove backend name from error string */
                $errstr = $abook->error;
                $errstr = ereg_replace('^\[.*\] *', '', $errstr);

                $formerror = $errstr;
                $orig_sel = $sel;
	    }
        }
	$showform = 'catmanager';
	break;

    case 'Rename Category':
	if (sizeof($sel) != 1) {
	    $formerror = 'Choose one category to rename';
            $orig_sel = $sel;
	}
	else {
	    $rename_from = $sel[0];
	}
	$showform = 'catmanager';
	break;

    case 'Set Category Name':
	$showform = 'catmanager';
	if ($newcat[newval] == 'Unfiled') {
	   $formerror = 'Invalid category name';
	   break;
	}
	$r = $abook->category_create($newcat[newval], $abook->localbackend);
	if (!$r) {
        	/* Remove backend name from error string */
                $errstr = $abook->error;
                $errstr = ereg_replace('^\[.*\] *', '', $errstr);

                $formerror = $errstr;
		break;
        }
	$r = $abook->category_move($newcat[newval], $newcat[oldval], $abook->localbackend);
	if (!$r) {
        	/* Remove backend name from error string */
                $errstr = $abook->error;
                $errstr = ereg_replace('^\[.*\] *', '', $errstr);

                $formerror = $errstr;
		break;
	}
	$subsel = array();
	array_push ($subsel, $newcat[oldval]);
	$r = $abook->category_drop($subsel, $abook->localbackend);
	if (!$r) {
        	/* Remove backend name from error string */
                $errstr = $abook->error;
                $errstr = ereg_replace('^\[.*\] *', '', $errstr);

                $formerror = $errstr;
        }
	break;

    case 'Go':
	if (!empty($prtfmt))
	    $showform = 'print';
	break;

    case 'Send selected':
        /************************************************
         * Send to address(es)                          *
         ************************************************/
        if (sizeof($sel) > 0) {
            $orig_sel = $sel;
            sort($sel);

            /* The selected addresses are identidied by "backend:uid". *
             * Sort the list and process one backend at the time            */
            $prevback  = -1;
            $subsel    = array();
	    $recipients = '';
            for ($i = 0 ; $i < sizeof($sel) ; $i++) {
                list($sbackend, $uid) = explode(':', $sel[$i]);

		$rec = $abook->lookup($uid, $sbackend);
		if (!empty( $rec['email'] ))
		    if (sizeof($sel) < 10)
			array_push($subsel, "{$rec['name']} <{$rec['email']}>");
		    else
			array_push($subsel, $rec[email]);

                /* When we get to a new backend, process addresses in *
                 * previous one.                                      */
                if ($prevback != $sbackend && $prevback != -1) {
		    $recipients .= implode(', ', $subsel);
                    $subsel   = array();
                }
                $prevback = $sbackend;
            }
	    $recipients .= implode(', ', $subsel);

	    /* Generate URL and pass to compose.php via http redirect */
	    $bcc_header = '';
	    $full_name = getPref($data_dir, $username, 'full_name');
	    if (sizeof($subsel) > 4)
		$bcc_header = rawurlencode("($full_name)")."&send_to_bcc=";
	    $url = 'compose.php?send_to='.$bcc_header.rawurlencode($recipients);
	    header('Location: ' . $url);
	    $abortform = true;
	    break;
	}
	break;

    case 'Upload File':
        /************************************************
         * Import a file                                *
         ************************************************/

	if (!is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	    $formerror = "You did not upload a file!";
	    unlink($_FILES['userfile']['tmp_name']);
	} else {
	    $error = false;
	    $formerror = '';
	    $fd = popen(SM_PATH . 'src/outlookimp ' . $_FILES['userfile']['tmp_name'] . ' ' . $username, 'r');

	    /*
	     * Look for Categories in table definition
	     */
	    $cats = array();
	    $tmpcat = 'Newly Imported';
	    while ($line = fgets($fd, 4096)) {
		if (strstr($line, 'Categories') && strlen($line) > 42)
		    $cats = explode("','", substr($line, 18, strlen($line)-42));
		if ($line == ");\n")
		    break;
	    }
	    array_push($cats, $tmpcat);
	    $newcat = array_diff($cats,$abook->category_get($abook->localbackend));
	    foreach ($newcat as $cat)
		$abook->category_create ($cat, $abook->localbackend);
		
	    $sqlcmd = '';
	    $uidlist = array();
	    $numrecs = 0;
	    while (($line = fgets($fd, 4096)) && !$error) {
		$sqlcmd .= $line;
//echo "Qline".substr($sqlcmd,strlen($sqlcmd)-3,3)."(".strlen($sqlcmd).")<br>";
		if (substr($sqlcmd, strlen($sqlcmd) - 3, 3) == ");\n") {
		    $ret = $abook->raw_sql_cmd($sqlcmd, $abook->localbackend);
		    if ($ret != true) {
			$error = true;
		    }
		    else {
			array_push($uidlist, substr(strstr($sqlcmd,"_"), 0, 10));
		        $numrecs++;
			$sqlcmd = '';
		    }
		}
	    }
	    pclose($fd);
	    unlink($_FILES['userfile']['tmp_name']);
	    if ($numrecs != 0) {
		$abook->category_addrecs ($tmpcat, $uidlist, $abook->localbackend);
		echo "<b><center>Notice: $numrecs records added to <i>$tmpcat</i> category successfully</center></b>";
		$active_cat = $tmpcat;
	    }
	    if ($error)
		$formerror = "Import: Bad data after $numrecs records";
	}
	$showform = 'addrlist';
	break;

    } /* switch ($action) */
} /* isset($command) */

    /***********************************************
     * Update/modify address                       *
     ***********************************************/
    if (!empty($editaddr) && $command != 'Cancel') {

		displayPageHeader($color, 'None');
		$page_header_sent = true;

                /* Stage one: Copy data into form */
                if (isset($sel) && sizeof($sel) > 0) {
                    if(sizeof($sel) > 1) {
                        $formerror = _("You can only edit one address at the time");
                        $showform = 'addrlist';
                        $defselected = $sel;
                    } else {
                        list($ebackend, $uid) = explode(':', $sel[0]);
                        $olddata = $abook->lookup($uid, $ebackend);

                        /* Display the "new address" form */
			$defdata = $olddata;
                        $showform = 'addrform';
                    }
                } else {

                    /* Stage two: Write new data */
                    if ($doedit = 1) {
                        $newdata = $editaddr;
 
			/* Collect forms data */
			if (sizeof($selcat) == 0)
			    $newdata['categories'] = '';
			else
			    $newdata['categories'] = implode(',', $selcat);

			/* Parse the city/state/ZIP */

			if (!parse_address ($newdata['homeaddress'], 'home', $newdata))
			    $formerror = _("Warning, did not recognize city/state/ZIP in home address");
			if (!parse_address ($newdata['businessaddress'], 'business', $newdata))
			    $formerror = _("Warning, did not recognize city/state/ZIP in business address");
			if (!parse_address ($newdata['otheraddress'], 'other', $newdata))
			    $formerror = _("Warning, did not recognize city/state/ZIP in other address");
			if (!empty( $newdata['homeaddress']) && empty( $newdata['businessaddress']))
			    $newdata['addresssel'] = 'H';

			for ($i = 1; $i <=5; $i++)
			    parse_phone ($newdata['phone'.$i], $i, $newdata);
			parse_date ($newdata['anniversary'], 'anniversary', $newdata);
			parse_date ($newdata['birthday'], 'birthday', $newdata);

                        $r = $abook->modify($oldnick, $newdata, $backend);

                        /* Handle error messages */
                        if (!$r) {
                            /* Display error */
                             echo html_tag( 'table',
                                html_tag( 'tr',
                                   html_tag( 'td',
                                      "\n". '<br><strong><font color="' . $color[2] .
                                      '">' . _("ERROR") . ': ' . $abook->error . '</font></strong>' ."\n",
                                      'center' )
                                   ),
                             'center', '', 'width="100%"' );

                            /* Display the "new address" form again */
			    $defdata = $newdata;
                        }
                    } else {

                        /* Should not get here... */
                        plain_error_message(_("Unknown error"), $color);
                        $abortform = true;
                    }
                }
    } /* !empty($editaddr)                  - Update/modify address */

    // Some times we end output before forms are printed
    if($abortform) {
       echo "</BODY></HTML>\n";
       exit();
    }
} /* if(sqgetGlobalVar('REQUEST_METHOD', $req_method, SQ_SERVER) */


/***************************************************
 * Finished with user commands, start page display *
 ***************************************************/

if (!$page_header_sent && $showform != 'print') {
   displayPageHeader($color, 'None');
   echo $javafunctions;
}

/* =================================================================== *
 * The following is only executed on a GET request, or on a POST when  *
 * a user is added, or when "delete" or "modify" was successful.       *
 * =================================================================== */

/* Display error messages */
if (!empty($formerror)) {
    echo html_tag( 'table',
        html_tag( 'tr',
            html_tag( 'td',
                   "\n". '<br><strong><font color="' . $color[2] .
                   '">' . _("ERROR") . ': ' . $formerror . '</font></strong>' ."\n",
            'center' )
        ),
    'center', '', 'width="100%"' );
}

/*
 * There are 3 different forms available:  the address list, the
 * address editing form, and the category manager.
 */

switch ($showform) {

/****************************
 * Display the address list *
 ****************************/
  case 'addrlist':

    /* Pick up selected category */
    if (isset($selcat))
	$active_cat = $selcat[0];

    /* Get and sort address list */
    $sortkey = get_sort($data_dir, $username);
    $abook->backends[$abook->localbackend]->set_sort($sortkey);
    list($sortkey, $sortdir) = sscanf ($sortkey, "%s %[A-Z]");
    $alist = $abook->list_addr();
    if(!is_array($alist)) {
        plain_error_message($abook->error, $color);
        exit;
    }
//    usort($alist,'alistcmp');

    $prevbackend = -1;
    $headerprinted = false;

    /* List addresses */
    if (count($alist) == 0) {
      plain_error_message('No matching addresses found', $color);
      echo '<FORM ACTION="' . $form_url . '" METHOD="POST">' . "\n";
      echo '<INPUT TYPE=submit NAME=command VALUE="' .
                                                     _("Create") . "\">\n";
      echo '</FORM>';
      break;
    }

    /* Show 25 rows unless specified in prefs or in URL */
    if (empty($list_numrows))
	$list_numrows = 25;
    if ($list_numrows > 0 && sizeof($alist) > $list_numrows) {

	/* Build the section links "Prev | Next | a b c | Show All" */
	$section = array();
	$section_uid = array();
	$section_start = 0;
        for ($i = 0; $i < sizeof($alist); $i += $list_numrows) {
	    /* Perform (slower) SQL query if non-standard columns specified */
	    if (sizeof(array_diff ($list_columns, array('Name', 'First Name',
		'Last Name', 'Company', 'Job Title', 'Email', 'Phone',
		'Notes'))) > 0)
		$row = $abook->lookup($alist[$i]['uid']);
	    else
		$row = $alist[$i];
	    array_push($section, substr(strtolower($row[$sortkey]),0,6));
	    array_push($section_uid, $row['uid']);
	    if (substr($start,0,1) != '_' &&
		strtolower($row[$sortkey]) < strtolower($start))
		$section_start = sizeof($section);
	    else if ($row['uid'] == $start)
		$section_start = sizeof($section) - 1;
	}
	$p1 = isset($active_cat) ? '&category='.rawurlencode($active_cat) : '';
	$p2 = '&list_numrows='.$list_numrows;
	$section_list  = "(Viewing $list_numrows of ".sizeof($alist).") ". (
			 $section_start > 0 ?
			 '<A HREF='.$form_url.'?start=' . 
			 $section_uid[$section_start-1].$p1.$p2.'>Prev</A> | ' :
			 'Prev | ' );
	$section_list .= $section_start < sizeof($section)-1 ?
			 '<A HREF='.$form_url.'?start='.
			 $section_uid[$section_start+1].$p1.$p2.'>Next</A> | ' :
			 'Next | ';
	$section_start = max(0,$section_start - 10);
	for ($i = $section_start;
             $i < min(sizeof($section), $section_start + 20); $i++) {
	    if ($i > 0 && empty($section[$i]))
		continue;
	    $section_list .= '<A HREF='.$form_url.'?start='.
			     $section[$i].$p1.$p2.'>'.(empty($section[$i]) ?
			      'First |' : substr(ucwords($section[$i]),0,2)) .
			     '</A> ';
	}
	$section_list .= '| <A HREF='.$form_url.'?list_numrows=-1'.$p1 .
			 '>Show All</A>';
    }

    echo '<FORM NAME=addr ACTION="' . $form_url . '" METHOD="POST">' . "\n";
    echo '<TABLE align=center border="0" cellpadding="1" cellspacing="0" width="90%">' . "\n";
                
    while(list($undef,$row) = each($alist)) {
    
	/* New table header for each backend */
	if($prevbackend != $row['backend']) {
                if($prevbackend < 0) {
		      echo
			   html_tag( 'tr' ) . "\n" .
			   html_tag( 'tr',
				 getSmallStringCell(_("Address..."), 'left', sizeof($list_columns)) .
				 html_tag ('td', print_picklist($printformats) .
				 '<INPUT TYPE=submit NAME=command VALUE="' .
					_("Go") . "\">\n",
					   'right', '', 'colspan="2"' ),
				 'nowrap', $color[9]) .
                           html_tag( 'tr',
                              html_tag( 'td',
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
                                              _("Create") . "\">\n" .
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
                                              _("Delete") . "\">\n" .
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
                                              _("Send selected") . "\">\n" .
                                   '<INPUT TYPE=button onClick="check_all(document.addr)" NAME=command VALUE="' .
                                              _("Select All") . "\">\n" .
			           categories_inp_field('',  'categories', '', 45, $active_cat, 1, $abook)  .
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
                                              _("Choose") . "\">\n" .
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
					      _("Insert") . "\">\n" .
                                   '<INPUT TYPE=submit NAME=command VALUE="' .
					      _("Remove") . "\">\n" ,
                                          'left', '', sprintf('colspan=%d',sizeof($list_columns)+2)),
			      '', $color[9]) .
			   (isset($section) ?
			   html_tag( 'tr',
			      html_tag( 'td', '<small>'.$section_list.'</small>',
					'center', '',
					sprintf('colspan=%d',sizeof($list_columns)+2)),
			           '', $color[9])
			   : '');
                }
    
		$td_str = '';
	        foreach( $list_columns as $colname ) {
		   if (substr($colname,0,4) == 'User')
		       $colname = $customlabels[substr($colname,strlen($colname)-1,1)];
		   $td_str .=
                         html_tag( 'th', _('&nbsp;'.$colname), 'left', '', 'width="1%"' );
		}
                echo  html_tag( 'tr', 
                                    html_tag( 'td', "\n" . '<strong>' . $row['source'] . '</strong>' . "\n", 'center', $color[0], sprintf('colspan=%d',sizeof($list_columns)+2) )
                                ) .
                      html_tag( 'tr', "\n" .
                          html_tag( 'th', '&nbsp;', 'left', '', 'width="1%"' ) .
                          html_tag( 'th', '&nbsp;', 'left', '', 'width="1%"' ) .
			  $td_str,
                      '', $color[9] ) . "\n";
    
                $line = 0;
                $headerprinted = true;
            } /* End of header */
    
            $prevbackend = $row['backend'];
    
	    /* Perform (slower) SQL query if non-standard columns specified */
	    if (sizeof(array_diff ($list_columns, array('Name', 'First Name',
		'Last Name', 'Company', 'Job Title', 'Email', 'Phone',
		'Notes'))) > 0)
		$row = $abook->lookup($row['uid']);

	    if (isset($start)) {
		if (substr($start,0,1) == '_' && $row['uid'] == $start)
		    unset($start);
		else if (!strncasecmp($start,$row[$sortkey],strlen($start)))
		    unset($start);
		else
		    continue;
	    }

            /* Check if this user is selected */
            if((in_array($row['backend'] . ':' . $row['uid'], $defselected)))
                $selected = 'CHECKED';
            else
                $selected = '';
    
            /* Print one row */
            $tr_bgcolor = '';
            if ($line % 2) { $tr_bgcolor = $color[0]; }
            if ($squirrelmail_language == 'ja_JP')
                {
            echo html_tag( 'tr', '') .
                html_tag( 'td',
                          '<SMALL>' .
                          '<INPUT TYPE=checkbox ' . $selected . ' NAME="sel[]" VALUE="' .
                          $row['backend'] . ':' . $row['uid'] . '"></SMALL>' ,
                          'center', '', 'valign="top" width="1%"' ) .
                html_tag( 'td', '&nbsp;' . $row['uid'] . '&nbsp;', 'left', '', 'valign="top" width="1%" nowrap' ) . 
                html_tag( 'td', '&nbsp;' . $row['LastName'] . ' ' . $row['FirstName'] . '&nbsp;', 'left', '', 'valign="top" width="1%" nowrap' ) .
                html_tag( 'td', '', 'left', '', 'valign="top" width="1%" nowrap' ) . '&nbsp;';
                } else {

            $td_str = html_tag( 'td',
                '<SMALL>' .
                '<INPUT TYPE=checkbox ' . $selected . ' NAME="sel[]" VALUE="' .
                $row['backend'] . ':' . $row['uid'] . '"></SMALL>' ,
                'center', '', 'valign="top" width="1%"' ) .

	    html_tag ( 'td', 
                '<A HREF="'.$form_url.'?edituid='."$row[backend]:$row[uid]".'"><IMG SRC="../images/editbutton.gif" height="12" width="17" ALT="[Edit]" border=0' .
		'</A>',
                'center', '', 'valign="top" width="1%"' );


	    foreach ($list_columns as $colname) {
		if ($colname == 'Email') {
	            $email = $abook->full_address($row);
	            if ($compose_new_win == '1') {
	                $td_1 = '<a href="javascript:void(0)" onclick=comp_in_new(false,"compose.php?send_to='.rawurlencode("$email ($row[name])").'")>';
	            }
	            else {
	                $td_1 = '<A HREF="compose.php?send_to=' .
			        rawurlencode("$email ($row[name])").'">';
	            }
	            $td_1 .= htmlspecialchars($row['email']) . '</A>&nbsp';
		    $td_str .= html_tag ('td', '&nbsp' . $td_1, 'left', '',
			'valign="top" width="5%" nowrap');
		}
		else {
		    if (!strcasecmp($colname, 'Phone'))
			$colname = 'phonedisplay';
		    $wrap = ($colname == 'Company') ? 'wrap' : 'nowrap';
		    $contents = $row[strtolower(str_replace(' ', '', $colname))];
		    if (!strcasecmp($colname, 'Birthday') ||
			!strcasecmp($colname, 'Anniversary'))
			$contents = date_fmt($contents, 'dd-mmm-yy');
		    if (!strcasecmp($colname, 'Name') &&
			$lastname_first)
			$contents = "{$row['lastname']}, {$row['firstname']}";
		    if (!strcasecmp($colname, 'Name') &&
		        !empty($highlight_cat) &&
			strstr(strtolower($row['categories']),$highlight_cat))
			$contents = "<b>$contents</b>";
		    $td_str .= html_tag ('td', '&nbsp;' . $contents .
			 '&nbsp;', 'left', '', 'valign="top" width="20%"'.$wrap );
	        }
	    }
	    echo  html_tag( 'tr', $td_str, '', $tr_bgcolor) . "\n";
            $line++;
	    if ($line == $list_numrows)
		break;
	}
    }
    
    /* End of list. Close table. */
    if ($headerprinted) {
        echo html_tag( 'tr',
                       html_tag( 'td',
                                '<INPUT TYPE="submit" NAME="command" VALUE="' . _("Send selected") .
                               "\">\n" ,
                       'center', '', 'colspan="5"' )
                     );
    }
    echo '</table></FORM>';

    break;


/*******************************************
 * Display the address form if edit/update *
 *******************************************/
case 'addrform':

    $button = ($action == 'Create') ? 'Add' : 'Update';

    /* Display the "new address" form */
    echo '<a name="AddAddress"></a>' . "\n" .
        '<FORM ACTION="' . $form_url . '" NAME=f_add METHOD="POST"' .
//	  ' onsubmit="return (
//   validateNumber(this.editaddr[phone1],
//      \'Please enter a phone number, numbers only\', 5, 10);"' .
	  ' ENCTYPE="multipart/form-data"' . ">\n" .
        html_tag( 'table',  
            html_tag( 'tr',
                html_tag( 'td', "\n". '<strong>' . sprintf(_("%s to %s"), $button, $abook->localbackendname) . '</strong>' . "\n",
                    'center', $color[0]
                )
            )
        , 'center', '', 'width="100%"' ) ."\n";

    address_form($abook, $button=='Add' ? 'addaddr' : 'editaddr', _($button . " address"), $defdata);

    if ($button == 'Update') {
         echo '<INPUT TYPE=hidden NAME=oldnick VALUE="' . $defdata['uid'] . '">' .
              '<INPUT TYPE=hidden NAME=backend VALUE="' .
              htmlspecialchars($defdata['backend']) . '">' .
              '<INPUT TYPE=hidden NAME=doedit VALUE=1>';
    }

    echo '</FORM>';
    break;

/*******************************************
 * Display the category manager            *
 *******************************************/
case 'catmanager':
    echo '<a name="Category Manager"></a>' . "\n" .
        '<FORM ACTION="' . $form_url . '" NAME=f_cat METHOD="POST">' . "\n" .
        html_tag( 'table',  
            html_tag( 'tr',
                html_tag( 'td', "\n". '<strong>' . sprintf(_("Categories in %s"), $abook->localbackendname) . '</strong>' . "\n",
                    'center', $color[0]
                )
            )
        , 'center', '', 'width="100%"' ) ."\n";

    $td_str = '<TABLE>';

    $r = $abook->category_get($abook->localbackend);
    sort ($r);
    foreach( $r as $j => $cat ) {
	if (isset($rename_from) && ($rename_from == $cat))
	     $left_col = '<INPUT NAME=newcat[newval] SIZE=20>';
	else
	     $left_col = '<INPUT TYPE=checkbox NAME="sel[]" VALUE="'.$cat.'">';
	$td_str .= html_tag( 'tr',
	     html_tag( 'td', 
                  $left_col, 'right', $color[0]) .
	     html_tag( 'td',
		  $cat.' <small>('.$abook->category_count($cat,
		  $abook->localbackend).' items)</small>', 'left', $color[0]),
	     'center', '', 'width="100%"' ) ."\n";
    }

    if (isset($rename_from)) {
	$left_col = '<INPUT TYPE=submit NAME=command VALUE="' . _("Set Category Name") . "\">\n" .
		    '<INPUT TYPE=hidden NAME=newcat[oldval] VALUE="' . $rename_from .  "\">\n";
    }
    else {
	$td_str .= html_tag( 'tr',
		  html_tag( 'td', '<INPUT NAME=newcat SIZE=20>', 'right', $color[0]) .
		  html_tag( 'td', '<INPUT TYPE=submit NAME=command VALUE="' . _("Add New Category") . "\">\n", 'left', $color[0]),
	     'center', '', 'width="100%"' ) ."\n";
	$left_col = '<INPUT TYPE=submit NAME=command VALUE="' . _("Delete Categories") . "\">\n" .
		    '<INPUT TYPE=submit NAME=command VALUE="' . _("Rename Category") . "\">\n";
    }
    $td_str .= html_tag ( 'tr',
		 html_tag ('td', $left_col,
			   'center', $color[4], 'colspan="2"' ),
	       'center', '', 'width="100%"') ."\n";

    echo $td_str . '</TABLE>';
    echo '</FORM>';
    break;

/****************************
 * Printable address list *
 ****************************/
//  case 'print-html':
//
//
//  This section was written to output address record details in html format.
//  It's scrapped because html doesn't support multicol except in old
//  Netscape browsers.  Code is kept because it may be recycle to display
//  a single record on screen.  See the Rich Text Format version below, which
//  generates a 3-column report similar to MS Outlook's detail report.
//
//    /* Pick up selected category */
//    if (isset($selcat))
//	$active_cat = $selcat[0];
//
//    /* Get and sort address list */
//    $alist = $abook->list_addr();
//    if(!is_array($alist)) {
//        plain_error_message($abook->error, $color);
//        exit;
//    }
//
//    usort($alist,'alistcmp');
//    $prevbackend = -1;
//
//    /* List addresses */
//    if (count($alist) == 0) {
//	plain_error_message('No matching addresses found', $color);
//	break;
//    }
//
//    $headerprinted = false;
//    while(list($undef,$row) = each($alist)) {
//    
//            /* New table header for each backend */
//            if($prevbackend != $row['backend']) {
//                if($prevbackend < 0) {
///*                    echo html_tag( 'table',
//			   html_tag( 'tr' ) . "\n",
//			    'center', $color[9], 'border="0" width="95%" cellpadding="1" cellspacing="0"');
//*/		}
//                $line = 0;
//echo '<div cols="3" gutter="15">';
//                $headerprinted = true;
//	    }
//            $prevbackend = $row['backend'];
//    
//            /* Print one row */
//            $tr_bgcolor = '';
//            if ($line % 2) { $tr_bgcolor = $color[0]; }
//
//	    $row = $abook->lookup($row[uid], $prevbackend);
//
//	    /* Build an array containing each row of the contact entry */
//	    $entry = array();
//	    array_push ($entry, $row[jobtitle]);
//	    if (!empty($row[company]))
//	      array_push ($entry, '<big>'.$row[company].'</big>');
//	    if (!empty($row[businessstreet])) {
//		array_push ($entry, $row[businessstreet]);
//		array_push ($entry, $row[businesscity].', '.$row[businessstate].' '.$row[businesszip].($row[businesscountry]==DEF_COUNTRY ? '' : ' '.$row[businesscountry]));
//	    }
//	    if (!empty($row[homestreet])) {
//		if (!empty($row[company]))
//		   array_push ($entry, '<b>Home:</b>');
//		array_push ($entry, $row[homestreet]);
//		array_push ($entry, $row[homecity].', '.$row[homestate].' '.$row[homezip].($row[homecountry]==DEF_COUNTRY ? '' : ' '.$row[homecountry]));
//	    }
//	    for ($i = 1; $i <=5 ; $i++)
//	      if (!empty($row['phone'.$i])) {
//		if ($row['phonesel'] == $i)
//		  array_push($entry, $row['phone'.$i].' <i>('.$phonelabels[$row['phone'.$i.'type']].')</i>');
//		else
//		  array_push($entry, $row['phone'.$i].' ('.$phonelabels[$row['phone'.$i.'type']].')');
//	      }
//	    for ($i = 1; $i <=3 ; $i++)
//	      if (!empty($row['email'.$i]))
//		  array_push($entry, $row['email'.$i]);
//
//	    if (strcmp($row[birthday],'0000-00-00'))
//		array_push($entry, 'Birthday: '.substr($row[birthday], 5, 5) .
//		   '-' . substr($row[birthday], 2, 2));
//	    if (strcmp($row[anniversary],'0000-00-00'))
//		array_push($entry, 'Anniversary: '.substr($row[anniversary], 5, 5) .
//		   '-' . substr($row[anniversary], 2, 2));
//	    if (!empty($row[sigother]))
//		array_push($entry, 'Sig Other: '.$row[sigother]);
//	    for ($i = 1; $i <= 4; $i++)
//		if (!empty($row['user'.$i]))
//		    array_push($entry, 'Custom '.$i.': '.$row['user'.$i]);
//	    array_push($entry, $row[categories]);
//	    if (!empty($row[notes]))
//	      array_push($entry, "<BR>".str_replace("\n", "<BR>", $row[notes]));
//
//	    $td_out = '';
//	    foreach( $entry as $j => $textline ) {
//		if (!empty($textline))
//		    $td_out .= $textline . '<BR>';
//	    }
//
///*	    $td_str = html_tag( 'tr',
//			html_tag ('td',
//			  '<b><font size="3">'.$row[name].'</font></b><br>',
//			   'center', $color[9], 'colspan="2"' ) .
//			html_tag ('td', $td_out,
//			   'center', $color[0], 'colspan="2"' ) .
//			'<p>',
//	       'center', '', 'width="50%"') ."\n";
//*/
//		$td_str = '<b><font size="+3">'.$row[name].'</font></b><br>' .
//			 '<font size="+2">'. $td_out .
//			'<p></font>' . "\n";
//
//	    echo $td_str;		      
//            $line++;
//    }
//    
//    echo '</div>';
//    echo '</table></FORM>';
//    break;

/****************************
 * Printable address list   *
 ****************************/
case 'print':

    /* Pick up selected category */
    if (isset($selcat))
	$active_cat = $selcat[0];

    if (sizeof($sel) == 0) {
        /* Get address list */
	$sortkey = get_sort();
	$abook->backends[$abook->localbackend]->set_sort($sortkey);
        $alist = $abook->list_addr();
	list($sortkey, $sortdir) = sscanf ($sortkey, "%s %[A-Z]");
        if(!is_array($alist)) {
	    plain_error_message($abook->error, $color);
	    exit;
	}
	/* List addresses */
	if (count($alist) == 0) {
	    plain_error_message('No matching addresses found', $color);
	    break;
	}

	$sel = array();
	foreach ($alist as $entry)
	   array_push ($sel, "$entry[backend]:$entry[uid]");
    }

    $prevbackend = -1;

/** preferences **/

    $prt_options = array();

    /*
     * Report preferences
     *   detail = 3col detail report
     *   env    = envelopes
     *   lb5160 = Avery 5160 labels
     *   wpages = White Pages style
     */

    $r = getPref($data_dir, $username, 'rpt_env_returnaddr');
    $prt_opt['rpt_env']['returnaddr'] = $r;

    $r = getPref($data_dir, $username, 'rpt_env_barcode');
    $prt_opt['rpt_env']['barcode'] = $r;

    $r = getPref($data_dir, $username, 'rpt_env_fontsize');
    $prt_opt['rpt_env']['fontsize'] = empty($r) ? 12 : $r;

    $r = getPref($data_dir, $username, 'rpt_env_fontname');
    $prt_opt['rpt_env']['fontname'] = empty($r) ? 'Arial' : $r;

    $r = getPref($data_dir, $username, 'rpt_lbl_barcode');
    $prt_opt['rpt_lb5160']['barcode'] = $r;

    $r = getPref($data_dir, $username, 'rpt_lbl_fontsize');
    $prt_opt['rpt_lb5160']['fontsize'] = empty($r) ? 10 : $r;

    $r = getPref($data_dir, $username, 'rpt_lbl_fontname');
    $prt_opt['rpt_lb5160']['fontname'] = empty($r) ? 'Arial' : $r;

    $prt_opt['rpt_wpages']['fontsize'] = 7;

    $r = getPref($data_dir, $username, 'rpt_wpages_sectmin');
    $prt_opt['rpt_wpages']['sectmin'] = empty($r) ? 150 : $r;

    $r = getPref($data_dir, $username, 'rpt_wpages_sectmin');
    $prt_opt['rpt_detail']['sectmin'] = empty($r) ? 50 : $r;

    /*
     * Form factor preferences
     *   pers  = 3.75" x 6.75"  (Filofax, DayRunner standard)
     *   half  = 5.5"  x 8.5"   (DayRunner)
     *   full  = 8.5"  x 11"    (Full page)
     */
    $prt_opt['form_pers']['fontname'] = 'Arial Narrow';
    $prt_opt['form_pers']['fontsize'] = 7;
    $prt_opt['form_half']['fontname'] = 'Arial Narrow';
    $prt_opt['form_half']['fontsize'] = 9;
    $prt_opt['form_full']['fontname'] = 'Arial';
    $prt_opt['form_full']['fontsize'] = 10;

    /*
     * General items
     */

    $prt_opt['general']['ret_addr'] = getPref($data_dir, $username, 'mailing_address');
    $prt_opt['general']['full_name'] = getPref($data_dir, $username, 'full_name');

    $r = getPref($data_dir, $username, 'rpt_gen_form_factor');
    $prt_opt['general']['form_factor'] = empty($r) ? 'full' : $r;

    $r = getPref($data_dir, $username, 'rpt_gen_sigother');
    $prt_opt['general']['sigother'] = $r;

    $r = getPref($data_dir, $username, 'rpt_gen_duplex');
    $prt_opt['general']['duplex'] = $r;

    $r = getPref($data_dir, $username, 'rpt_gen_suppress_comma');
    $prt_opt['general']['suppress_comma'] = $r;

/* end prefs */

    /* Generate raw file header of type RTF */
    $now = gmdate('D, d M Y H:i:s') . ' GMT';
    if (strstr (strtolower($prtfmt), 'vcard')) {
	$ext       = 'vcf';
	$mime_type = 'text/x-vcard; charset=us-ascii';
	$filename  = "Contacts ($active_cat)";
    }
    else if (strstr (strtolower($prtfmt), 'csv')) {
	$ext       = 'csv';
	$mime_type = 'text/x-csv; charset=us-ascii';
	$filename  = "Contacts ($active_cat)";
    }
    else {
	$ext       = 'rtf';
        $mime_type = 'application/rtf';
	$filename  = "Contacts ($active_cat-$prtfmt)";
    }
    header('Content-Type: ' . $mime_type);
    header('Expires: ' . $now);
    header('Content-Disposition: attachment; filename="' . $filename . '.' . $ext . '"');
//    header('Pragma: no-cache');
//    header('Cache-control: private');
    $section = '';

    switch ($prtfmt) {
	/*
	 * 3-column detail report
	 */
	case "print $printformats[0]":

	    /* Generate alphabetic breaks if more than 50 (or as otherwise
	     * configured) names printed */
	    $section_breaks = (sizeof($sel) > $prt_opt[rpt_detail][sectmin]) ?
			 true : false;

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);
    
	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
		       echo rtf_header($prt_opt[general][form_factor] == 'full');
		       echo rtf_page_format($prt_opt[general][form_factor], $prt_opt);
		       echo '\fs20 ';
		    }
		}
	        $prevbackend = $sbackend;

        /* Retrieve all fields of the current record */

	$row = $abook->lookup($uid, $sbackend);
    
	if ($section_breaks &&
		  strcmp(strtolower(substr($row[$sortkey],0,1)), $section)) {
	    $section = strtolower(substr($row[$sortkey],0,1));
	    $alpha_header = '{\b\fs28\chshdng0\chcfpat0\chcbpat1 | '.
				$section.' |}'."\n".'\par\par ';
	}
	else
	    $alpha_header = '';


	/* Build an array containing each row of the contact entry */
	$entry = array();
	array_push ($entry, $row[jobtitle]);
	if (!empty($row[company]))
	    array_push ($entry, '{\fs24 '.$row[company].'}');
	if (!empty($row[businessaddress])) {
	    array_push ($entry, str_replace("\n", "\n\\par ",
			    $row[businessaddress]));
	    if (!empty($row[businesscountry]) && $row[businesscountry] != DEF_COUNTRY)
		   array_push ($entry, $row[businesscountry]);
	}
	if (!empty($row[homeaddress])) {
	    if (!empty($row[company]) || !empty($row[businessaddress]))
		   array_push ($entry, '{\b Home: }');
	    array_push ($entry, str_replace("\n", "\n\\par ",
			    $row[homeaddress]));
	    if (!empty($row[homecountry]) && $row[homecountry] != DEF_COUNTRY)
		   array_push ($entry, $row[homecountry]);
	}
	if (!empty($row[otheraddress])) {
	    array_push ($entry, '{\b Other: }');
	    array_push ($entry, str_replace("\n", "\n\\par ",
			    $row[otheraddress]));
	    if (!empty($row[othercountry]) && $row[othercountry] != DEF_COUNTRY)
		array_push ($entry, $row[othercountry]);

	}
	for ($i = 1; $i <=5 ; $i++)
	    if (!empty($row['phone'.$i])) {
		if ($row['phonesel'] == $i)
		  array_push($entry, $row['phone'.$i].' {\i ('.
			     $phonelabels[$row['phone'.$i.'type']].')}');
		else
		  array_push($entry, $row['phone'.$i].' ('.
			     $phonelabels[$row['phone'.$i.'type']].')');
	    }
	for ($i = 1; $i <=3 ; $i++)
	    if (!empty($row['email'.$i]))
		array_push($entry, $row['email'.$i]);

	if ($row[birthday] != '0000-00-00')
		array_push($entry, 'Birthday: '.date_fmt($row[birthday],'mm-dd-yy'));
	if ($row[anniversary] != '0000-00-00')
		array_push($entry, 'Anniversary: '.date_fmt($row[anniversary],'mm-dd-yy'));
	if (!empty($row[sigother]))
		array_push($entry, 'Sig Other: '.$row[sigother]);
	if (!empty($row[children]))
		array_push($entry, 'Children: '.$row[children]);
	if (!empty($row[webpage]))
		array_push($entry, 'Web Page: '.$row[webpage]);
	for ($i = 1; $i <= 4; $i++)
		if (!empty($row['user'.$i]))
		    array_push($entry, $customlabels[$i].': '.$row['user'.$i]);
	if (!empty($row[categories]))
	       array_push($entry, '{\fs18 '.$row[categories].'}');
	if (!empty($row[notes]))
	      array_push($entry, '\par '.str_replace("\n", "\n\\par ", trim($row[notes])));

	$td_out = '';
	foreach( $entry as $j => $textline ) {
	    if (!empty($textline))
		$td_out .= $textline . "\n\\par ";
	}

	$td_str = '\pard \keepn\widctlpar\adjustright'.$alpha_header .
		  '\shading2500\cbpat8{\b\fs24 '.$row[name]."\n" .
		  '\par }\pard \keepn\widctlpar\adjustright {' .
		  $td_out .
		  '}\pard \widctlpar\adjustright' . "\n" . '\par ';

	echo $td_str;		      
    }
    
	    echo '}}';
	    break;

	/*
	 * Telephone directory style
	 */
	case "print $printformats[1]":

	    /* Generate alphabetic breaks if more than 150 (or as otherwise
	     * configured) names printed */
	    $section_breaks = (sizeof($sel) > $prt_opt[rpt_wpages][sectmin]) ?
			 true : false;


            if ($prt_opt[general][form_factor] == 'pers') {
		  $txpos = 4100;
		  $tpos  = '\shpleft-461\shptop8208\shpright-173\shpbottom9360';
		  $fs    = 7;
	    }
	    else if ($prt_opt[general][form_factor] == 'half') {
		  $txpos = 3130;
		  $tpos  = '\shpleft-576\shptop10656\shpright-288\shpbottom11808';
		  $fs    = 9;
	    }
	    else {
		  $txpos = 3240;
		  $fs    = 9;
		  $tpos  = 0;
	    }
	    $ftr = !$tpos ? '' : '{\lang1024
{\shp{\*\shpinst'.$tpos.'\shpfhdr0\shpbxcolumn\shpbypage\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}
{\sp{\sn lTxid}{\sv 65536}}{\sp{\sn dxTextLeft}{\sv 64008}}{\sp{\sn dyTextTop}{\sv 0}}{\sp{\sn dxTextRight}{\sv 18288}}{\sp{\sn txflTextFlow}{\sv 2}}{\sp{\sn fLine}{\sv 0}}{\shptxt \pard\plain \widctlpar\adjustright \f16\fs18\cgrid {\fs'.intval($fs*4/3).'\f40
Page printed {\field{\*\fldinst DATE \\\\@ "M/d/yy" \\\\* MERGEFORMAT }}{\fldrslt }
\par }}}{\shprslt{\*\do\dobxcolumn\dobypage\dodhgt8192\dptxbx{\dptxbxtext\pard\plain \widctlpar\adjustright \f16\fs18\cgrid {\fs12 Page
\par }}\dpx-576\dpy10800\dpxsize288\dpysize1008\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr255\dpfillbgcg255\dpfillbgcb255\dpfillpat1\dplinehollow}}}}';

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);
    
	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
		       echo rtf_header($prt_opt[general][form_factor] == 'full');
		       echo rtf_page_format($prt_opt[general][form_factor], $prt_opt);
		       echo '\pard\plain \fi-180\li180\ri1152\widctlpar\tqr\tx'.$txpos.'\adjustright \f1\fs14\cgrid';
		       echo $ftr;
		    }
		}
	        $prevbackend = $sbackend;
    
	        /* Retrieve all fields of the current record */
		$row = $abook->lookup($uid, $sbackend);

		if (empty($row['phonedisplay']))
		    continue;

		switch ( $row['addresssel'] ) {
		    case 'B':
		        $mailingaddress = (empty($row['company']) ?
			   $row['jobtitle'] : $row['company']);
			if (!empty($mailingaddress) && !empty($row['businesscity']))
			   $mailingaddress .= "\n\\line ";
			$mailingaddress .= (empty($row['businessstreet']) ? '' :
			   str_replace("\n", ", ", $row['businessstreet']).", ") .
			   "{$row['businesscity']} " .
			   substr($row['businesszip'],0,5);
			break;
		    case 'O':
		        $mailingaddress = (empty($row['otherstreet']) ? '' :
			   str_replace("\n", ", ", $row['otherstreet']).", ") .
			   "{$row['othercity']} " . substr($row['otherzip'],0,5);
			break;
		    case 'H':
		    default:
		        $mailingaddress = (empty($row['homestreet']) ? '' :
			   str_replace("\n", ", ", $row['homestreet']).", ") .
			   "{$row['homecity']} " . substr($row['homezip'],0,5);
		}
		$mailingaddress = trim($mailingaddress);

		if ($section_breaks &&
		      strtolower(substr($row[$sortkey],0,1)) != $section) {
		    $section = strtolower(substr($row[$sortkey],0,1));
		    $alpha_header = '\pard \qc\fi-180\li180\ri24\keepn\widctlpar\tqr\tx3240\adjustright \cbpat1\ {\b\caps\fs18 '. $section . "\n" .
			'\par }\pard \fi-180\li180\ri1152\keep\widctlpar\tqr\tx'.$txpos.'\adjustright ';
		}
		else
		    $alpha_header = '';

		$td_out = "\keep {\\b\\caps $row[lastname]} {\\b $row[firstname]} ";
		$td_out .= "{\\fs12 $mailingaddress}{\\uld \\tab}";
		$td_out .= substr($row['phonedisplay'],strlen($row['phonedisplay'])-3,3).
			'{\b '.substr($row['phonedisplay'],0,strlen($row['phonedisplay'])-3).'}'."\n";
		for ($i = 1; $i <= 5; $i++) {
		    if (!empty($row['phone'.$i]) && $row[phonesel] != $i) {
			$td_out .= '\line {\uld\tab}('.substr($phonelabels[$row['phone'.$i.'type']],0,1).
			 '){\b '.$row['phone'.$i].'}' . "\n";
		    }
		}
		$td_out .= '\par ';
		echo $alpha_header . $td_out;
	    }
	    echo '}}';

	    break;

	/*
	 * Table format
	 */
	case "print $printformats[2]":

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
			echo rtf_header_table(false);
			echo '{Name\cell Company\cell }\pard \qc\widctlpar\intbl\adjustright {Work Phone\cell Email\cell Home Phone\cell Cell Phone\cell Fax\cell Birthday\cell Sig Other\cell }';
			echo "\n\\par Hello, world" . "\n\\par ";
		    }
		}
	        $prevbackend = $sbackend;
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		$td_out = '\trowd \trgaph72\trleft-72\trbrdrt\brdrs\brdrw10 \trbrdrl\brdrs\brdrw10 \trbrdrb\brdrs\brdrw10 \trbrdrr\brdrs\brdrw10 \trbrdrh\brdrs\brdrw10 \trbrdrv\brdrs\brdrw10 
\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx1980\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx3780\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 
\clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx5220\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx7074\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx8514\clvertalt
\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx9954\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx11394\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr
\brdrs\brdrw10 \cltxlrtb \cellx12540\clvertalt\clbrdrt\brdrs\brdrw10 \clbrdrl\brdrs\brdrw10 \clbrdrr\brdrs\brdrw10 \cltxlrtb \cellx14118\pard \widctlpar\intbl\adjustright ';

		$td_out .= '{'.$row['name'].'}{';
		$td_out .= '\cell {'.$row['company'].'}\cell {'.$row['workphone'].'}' .
			   '\cell {'.$row['email'].'}\cell {'.$row['homephone'].'}' .
			   '\cell {'.$row['cellphone'].'}\cell {'.$row['fax'].'}';
		$td_out .= '\cell {'.$row['birthday'].'}';
		$td_out .= '\cell {'.$row['sigother'].'}}';
			   
		echo $td_out;
	    }
	    echo '\pard \widctlpar\intbl\adjustright {\row }\pard \widctlpar\adjustright {
\par } ';
	    echo '}}';
	    break;

	/*
	 * Envelopes: Number 10 or 6-3/4
	 */
	case "print $printformats[3]":
	case "print $printformats[4]":

	    $pg_break = '';
	    $sigother_prune = array();

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
			echo rtf_header(false);
		        echo '{\fonttbl{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$prt_opt['rpt_env']['fontname'].';}}';
			echo ($prtfmt == 'print Envelopes #10') ?
				'\paperw13680\paperh5940' :
				'\paperw9360\paperh5220';
			echo '\margl576\margr720\margt360\margb720 \widowctrl\ftnbj\aenddoc\pgnstart0\hyphcaps0\formshade\viewkind1\viewscale129\viewzk2\pgbrdrhead\pgbrdrfoot \fet0';
		    }
		}
	        $prevbackend = $sbackend;
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		/* Skip record if sigother already printed */
		if ($prt_opt['general']['sigother'] &&
		    array_search($row['name'], $sigother_prune))
		    continue;

		/* Skip record if no street address */
		switch ($row['addresssel']) {
		    case 'B':
			$street = $row['businessstreet']; break;
		    case 'O':
			$street = $row['otherstreet']; break;
		    default:
		    case 'H':
			$street = $row['homestreet'];
		}
		if (empty($street))
		    continue;

		$td_out = $pg_break . '\sectd\lndscpsxn\binfsxn4\pgnrestart\pgnstarts0\linex0\endnhere\sectdefaultcl\pard\plain \s16\widctlpar\adjustright \f1\fs'.($prt_opt['rpt_env']['fontsize']*2-4).'\cgrid ';
		$pg_break = '\par {\page }';
		if ($prt_opt['rpt_env']['returnaddr'])
		  $td_out .= '{'.$prt_opt['general']['full_name']."\n\\par " .
			     str_replace("\n", "\n\\par ",
			     $prt_opt['general']['ret_addr']) .
			     "\n\\par }";
		else
		  $td_out .= "\n\\par ";
		
		$td_out .= '\pard\plain \s15\li2880\widctlpar\phpg\posxc\posyb\absh-1980\absw7920\dxfrtext180\dfrmtxtx180\dfrmtxty0\adjustright \f1\fs'.($prt_opt['rpt_env']['fontsize']*2).'\cgrid ';

		switch ($row['addresssel']) {
		    case 'B':
			$company = $row['company'];
			$country = $row['businesscountry'];
			$zip = $row['businesszip']; break;
		    case 'O':
			$company = $row['company'];
			$country = $row['othercountry'];
			$zip = $row['otherzip'];    break;
		    case 'H':
		    default:
			$company = '';
			$country = $row['homecountry'];
			$zip = $row['homezip'];
		}
		if ($prt_opt['rpt_env']['barcode'] && !empty($zip) &&
		    country_is_us($country))
		    $td_out .= '{\field{\*\fldinst {\lang1024 BARCODE \\\\u "'.
		       $zip.'" \\\\* MERGEFORMAT }}{\fldrslt }}{\lang1024'.
		       "\n\\par }";
		$td_out .= '{'.$row['name']."\n\\par ";
		if ($prt_opt['general']['sigother'] && !empty($row['sigother'])) {
		    $td_out .= "$row[sigother]\n\\par ";
		    array_push ($sigother_prune, $row['sigother']);
		}
		$mailingaddress = $row['mailingaddress'];
		if ($prt_opt['general']['suppress_comma'])
		    $mailingaddress = substr_replace ($mailingaddress, '', strrpos($mailingaddress, ','), 1);
		$td_out .= $company . (!empty($company) ? "\\n\par " : '') .
		   str_replace("\n", "\n\\par ", $mailingaddress) .
		   (country_is_us($country) ? '' : "\n\\par " . $country) . "}\n";
		echo $td_out;
	    }
	    echo '}}';
	    break;

	/*
	 * Business cards (2" x 3.5", 2 columns of 5 per page)
	 * Rolodex cards (2.25" x 4", 2 columns of 4 per page)
	 */
	case "print $printformats[5]":
	case "print $printformats[7]":
	    $rolo = strstr($prtfmt, 'Rolo') ? true : false;
	    $wd =   ($rolo ? 2700 : 2520);
	    $ht =   ($rolo ? 3168 : 2880);
	    $indt = ($rolo ? 252  : 126);
	    $xpos = array ($wd-15, $wd*2-15, $wd*3-15, $wd*4-15);
	    $celldef = '\trowd \trgaph15\trrh-'.$ht.'\trleft-15\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth'.$wd.' \cellx'.$xpos[0].'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth'.$wd.' \cellx'.$xpos[1].'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth'.$wd.' \cellx'.$xpos[2].'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth'.$wd.' \cellx'.$xpos[3];

	    $col = 0;

	    echo rtf_header_labels(false);
	    echo $rolo ? '\margl720\margr720\margt1530\margb1620 ' :
		         '\margl1080\margr1080\margb720 ';
	    echo $celldef;
	    echo  '\pard\plain \ql \li'.$indt.'\ri'.$indt.'\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin'.$indt.'\lin'.$indt.' 
\fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033 {';

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		$td_out = '';
		if ($col++ == 2) {
		    $td_out = '}\pard \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {' . $celldef . '\row }\pard \ql \li'.$indt.'\ri'.$indt.'\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin'.$indt.'\lin'.$indt.' {';
		    $col = 1;
		}
		$td_out .= '\pard \li'.$indt.'\ri'.$indt.'\widctlpar\intbl\adjustright ';
		if (!empty($row['company'])) {
		    $td_out .= '\f1\cbpat3 {\fs22\b '.$row['company']."}\n\\par ";
		    $td_out .= '\pard \li'.$indt.'\ri'.$indt.'\widctlpar\intbl\adjustright ';
		}
		if (!empty($row['businessaddress']))
		    $td_out .= "\n\\par\n\\par\n\\par {\fs16 " .
			str_replace("\n", "\n\\par ", $row['businessaddress']) . "}\n";
		$td_out .= '\cell \pard \qr\li'.$indt.'\ri'.$indt.'\widctlpar\intbl\adjustright ';
		$td_out .= '\f1\fs26 ' . $row['name'] . "\n\\par ";
		if (!empty($row['jobtitle']))
		    $td_out .= '{\fs22\i ' . $row['jobtitle'] . "}\n\\par ";
		$td_out .= '\fs16 ';
		for ($i = 1; $i <= 5; $i++)
		    if (!empty($row['phone'.$i])) {
			switch ($row['phone'.$i.'type']) {
			    case '0':  $td_out .= 'Tel. '; break;
			    default:
			      $td_out .= $phonelabels[$row['phone'.$i.'type']] .
				' ';
			}
		        $td_out .= $row['phone'.$i] . "\n\\par ";
		    }
		if (!empty($row['email']))
		    $td_out .= $row['email'] . "\n\\par ";
		if (!empty($row['homeaddress']))
		    $td_out .= "\n\\par\n\\par " .
			str_replace("\n", "\n\\par ", $row['homeaddress']) . "\n";
		$td_out .= '\cell ';

		echo $td_out;
		
	    }
	    echo '}\pard \widctlpar\intbl\adjustright {\row }\pard \widctlpar\adjustright {\v 
\par }}';
	    break;

	/*
	 * Mailing labels (Avery 5160 1" x 2-5/8", 3 columns of 10 per page)
	 */
	case "print $printformats[6]":

	    $sigother_prune = array();

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
		       echo rtf_header_labels(false);
		       echo '{\fonttbl{\f40\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$prt_opt['rpt_lb5160']['fontname'].';}}';

		       echo '\trowd \trgaph15\trrh-1440\trleft-15\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 \clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx3765\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth180 \cellx3945\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx7725\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth180 \cellx7905\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone 
\clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx11685\pard\plain \ql \li95\ri95\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin95\lin95 \fs24\lang1033\langfe1033\cgrid\langnp1033\langfenp1033\f40\fs'.($prt_opt['rpt_lb5160']['fontsize']*2).' {';

			$col = 0;
		    }
		}
	        $prevbackend = $sbackend;
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		/* Skip record if sigother already printed */
		if ($prt_opt['general']['sigother'] &&
		    array_search($row['name'], $sigother_prune))
		    continue;

		/* Skip record if no street address */
		switch ($row['addresssel']) {
		    case 'B':
			$street = $row['businessstreet']; break;
		    case 'O':
			$street = $row['otherstreet']; break;
		    default:
		    case 'H':
			$street = $row['homestreet'];
		}
		if (empty($street))
		    continue;
		$td_out = '';
		if ($col++ == 3) {
		    $td_out = '}\pard \ql \li0\ri0\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin0\lin0 {\trowd \trgaph15\trrh-1440\trleft-15\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 \clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone 
\clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx3765\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth180 \cellx3945\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone 
\clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx7725\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth180 \cellx7905\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone 
\clbrdrb\brdrnone \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3780 \cellx11685\row }\pard \ql \li95\ri95\widctlpar\intbl\aspalpha\aspnum\faauto\adjustright\rin95\lin95 {';
		    $first_break = '';
		    $col = 1;
		}
		switch ($row['addresssel']) {
		    case 'B':
			$company = $row['company'];
			$country = $row['businesscountry'];
			$zip = substr($row['businesszip'],0,5); break;
		    case 'O':
			$company = $row['company'];
			$country = $row['othercountry'];
			$zip = substr($row['otherzip'],0,5);    break;
		    case 'H':
		    default:
			$company = '';
			$country = $row['homecountry'];
			$zip = substr($row['homezip'],0,5);
		}
		if ($prt_opt['rpt_lb5160']['barcode'] && !empty($zip) &&
		    country_is_us($country))
		    $td_out .= '{\field{\*\fldinst {\lang1024 BARCODE \\\\u "'.
		       $zip.'" \\\\* MERGEFORMAT }}{\fldrslt }}{\lang1024'.
		       "\n\\par }";
		$td_out .= $row['name']."\n\\par ";
		if ($prt_opt['general']['sigother'] && !empty($row['sigother'])) {
		    $td_out .= "$row[sigother]\n\\par ";
		    array_push ($sigother_prune, $row['sigother']);
		}
		$td_out .= $company . (!empty($company) ? "\\n\par " : '');
		$mailingaddress = $row['mailingaddress'];
		if ($prt_opt['general']['suppress_comma'])
		    $mailingaddress = substr_replace ($mailingaddress, '', strrpos($mailingaddress, ','), 1);
		$td_out .= str_replace("\n", "\n\\par ", $mailingaddress);
		if (!country_is_us($country))
		    $td_out .= "\n\\par " . $country;
		$td_out .= '\cell ' . ($col < 3 ? '\cell ' : '');

		echo $td_out;
	    }
	    if ($col == 2) 
		echo '\cell ';
	    else if ($col == 1)
		echo '\cell \cell \cell';
	    echo '}\pard \widctlpar\intbl\adjustright {\row }\pard \widctlpar\adjustright {\v 
\par }}';
	    break;
	
	/*
	 * Fax cover sheet
	 */
	case "print $printformats[8]":
	    $first_page = true;
	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* New table header for each backend */
	        if($prevbackend != $sbackend) {
	            if($prevbackend < 0) {
		       echo rtf_header(false);
		       echo '\margl1008\margr1008 ';
		       echo '{\fonttbl{\f1\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$prt_opt[form_full][fontname].';}}';
		    }
		}
	        $prevbackend = $sbackend;
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		$td_out = '';
		if (!$first_page)
		    $td_out .= '\pard \sa240\widctlpar\page ' . "\n\\par ";
		$first_page = false;
		$td_out .= '\pard\plain \cols1\s30\sl200\slmult0
\keep\widctlpar\pvpg\phpg\posx1011\posy1008\absh-1148\absw5040\abslock1\dxfrtext187\dfrmtxtx187\dfrmtxty187\nowrap\adjustright \f1\fs'.(($prt_opt[form_full][fontsize]+1)*2).'\expndtw-2\cgrid {';

		/*
		 * Text box with white-on-black "Fax"
		 */
		$td_out .= '{\shp{\*\shpinst\shpleft8488\shptop-144\shpright10360\shpbottom720\shpfhdr0\shpbxcolumn\shpbypara\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}
{\sp{\sn lTxid}{\sv 65536}}{\sp{\sn hspNext}{\sv 1026}}{\sp{\sn fLine}{\sv 0}}{\shptxt \pard\plain \qc\widctlpar\adjustright \cbpat1 \f1\fs20\cgrid {\fs52\b Fax
\par }}}{\shprslt{\*\do\dobxcolumn\dobypara\dodhgt8192\dptxbx{\dptxbxtext\pard\plain \qc\widctlpar\adjustright \cbpat1 \f1\fs20\cgrid {\fs52\b Fax
\par }}\dpx7488\dpy-144\dpxsize1872\dpysize864\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr255\dpfillbgcg255\dpfillbgcb255\dpfillpat1\dplinehollow}}}';

		/*
		 * Return address
		 */
		$td_out .= '{\fs'.(($prt_opt[form_full][fontsize]+1)*2).' '.str_replace("\n", "\n\\par ", getPref($data_dir, $username, 'mailing_address'))."\n\\par ".getPref($data_dir, $username, 'email_address')."\n\\par }";
		$td_out .= '\pard \sa240\widctlpar\brdrt\brdrs\brdrw10\brsp20 \brdrb\brdrs\brdrw10\br\sp20 \tx900\tx4590\tx5580\adjustright {';
		$td_out .= '{\b To:}\tab ' . $row['name'] . '\tab ';
		$td_out .= '{\b From:}\tab ' . $prt_opt['general']['full_name'];
		if (!empty($row['company']))
		    $td_out .= '\line\tab\sa0 {\i ' . $row['company'] . '}';
		$td_out .= "\n\\par\\sa240 ";
		$td_out .= '{\b Fax:}\tab ' . $row['fax'] . '\tab {\b Pages:}' .
				"\n\\par ";
		$td_out .= '{\b Phone:}\tab ' . $row['phone'.$row['phonesel']] . '\tab ';
		$td_out .= '{\b Date:}\tab ' . '{\field{\*\fldinst DATE \\\\@ "MMMM d, yyyy" \\\\* MERGEFORMAT }}{\fldrslt }' . "\n\\par ";
		$td_out .= '{\b Subject:}\tab \tab {\b Cc:}' . "\n\\par ";
		$td_out .= '\pard \sa0\widctlpar\br\sp20 \tx810\tx4590\tx5580\adjustright 
\par ';
		$td_out .= '}}';

		echo $td_out;
	    }
	    echo '}}';
	    break;

	/*
	 * DayRunner 8.5x5.5 pages
	 * DayRunner 6.75x3.75 pages
	 */
	case "print $printformats[9]":
	case "print $printformats[10]":

	    $sect_tabs = array ( 'A ~ B', 'C ~ D', 'E ~ F', 'G ~ H', 'I ~ J',
			   'K ~ L', 'M ~ N', 'O ~ Q', 'R ~ S', 'T ~ U',
			   'V ~ W', 'X ~ Z', '~');
	    if (strstr ($prtfmt, '5.5')) {
		$pgwd = 550;		// Page width 5.50"
		$pght = 850;		// Page height 8.50"
		$ht   = 304;  		// Row height (15.2 points)
		$ofst = 504;		// Hole-punch offset (.35")
		$font = 'Arial Narrow';
		$fs   = 9; 		// Font size
		$cl1  = 3348;
		$cl2  = 6804;
		$prpg = 6;		// 6 addresses per page
		$tpos = '\shpleft-576\shptop10656\shpright-288\shpbottom11808';
	    }
	    else {
		$pgwd = 375;		// Page width 3.75"
		$pght = 675;		// Page height 6.75"
		$ht   = 281;  		// Row height (14.05 points)
		$ofst = 360;		// Hole-punch offset (.25")
		$font = 'Arial Narrow';
		$fs   = 7; 		// Font size
		$cl1  = 2169;
		$cl2  = 4446;
		$prpg = 5;		// 5 addresses per page.
		$tpos = '\shpleft-461\shptop8208\shpright-173\shpbottom9360';
	    }

	    echo rtf_header_labels(false);
	    echo '{\fonttbl{\f40\fswiss\fcharset0\fprq2{\*\panose 020b0506020202030204}'.$font.';}}';
	    echo '\paperw'.($pgwd*144/10).'\paperh'.$pght*144/10;
	    if ($prt_opt['general']['duplex'])
		echo '\margmirror ';
	    echo '\margl360\margr360\margt360\margb360\gutter'.$ofst.'\f40\fs' . $fs*2;

	    $celldef = '\trowd \trgaph108\trrh-HT\trleft-108\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 
\trbrdrt\brdrtnthsg\brdrw60 \trbrdrl\brdrtnthsg\brdrw60
\trbrdrb\brdrthtnsg\brdrw60 \trbrdrr\brdrthtnsg\brdrw60
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrhair \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3363 \cellx'.$cl1.'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrhair \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3363 \cellx'.$cl2;
	    $hdrdef = '\trowd \trgaph108\trrh-400\trleft-108\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 
\trbrdrt\brdrtnthsg\brdrw60 \trbrdrl\brdrtnthsg\brdrw60
\trbrdrb\brdrthtnsg\brdrw60 \trbrdrr\brdrthtnsg\brdrw60
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth6819\clcbpat3\clshdng2500 \cellx'.$cl2;

	    $cell2def = '\trowd \trgaph108\trrh-HT\trleft-108\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 
\trbrdrt\brdrtnthsg\brdrw60 \trbrdrl\brdrtnthsg\brdrw60
\trbrdrb\brdrthtnsg\brdrw60 \trbrdrr\brdrthtnsg\brdrw60
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrhair \clbrdrr\brdrhair \cltxlrtb\clftsWidth3\clwWidth3363\clcbpat8\clshdng2000 \cellx'.$cl1.'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrhair \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3363\clcbpat8\clshdng2000 \cellx'.$cl2;
	    $cell3def = '\trowd \trgaph108\trrh-HT\trleft-108\trkeep\trftsWidth1\trpaddl15\trpaddr15\trpaddfl3\trpaddfr3 
\trbrdrt\brdrtnthsg\brdrw60 \trbrdrl\brdrtnthsg\brdrw60
\trbrdrb\brdrthtnsg\brdrw60 \trbrdrr\brdrthtnsg\brdrw60
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs \clbrdrr\brdrhair \cltxlrtb\clftsWidth3\clwWidth3363\clcbpat8\clshdng2000 \cellx'.$cl1.'
\clvertalc\clbrdrt\brdrnone \clbrdrl\brdrnone \clbrdrb\brdrs \clbrdrr\brdrnone \cltxlrtb\clftsWidth3\clwWidth3363\clcbpat8\clshdng2000 \cellx'.$cl2;
	    $newrow = str_replace("HT", $ht, $celldef) . '\pard \ql \widctlpar\intbl\aspalpha\aspnum\faauto\adjustright {\pard \widctlpar\intbl\adjustright ';
	    $newrow2 = str_replace("HT", $ht, $cell2def) . '\pard \ql \widctlpar\intbl\aspalpha\aspnum\faauto\adjustright {\pard \widctlpar\intbl\tx630\adjustright ';
	    $newrow3 = str_replace("HT", $ht, $cell3def) . '\pard \ql \widctlpar\intbl\aspalpha\aspnum\faauto\adjustright {\pard \widctlpar\intbl\tx630\adjustright ';
	    $endrow = '}\pard \ql \widctlpar\intbl\aspalpha\aspnum\faauto\adjustright {\row }';

	    $rightcol = '\cell \pard \ql\widctlpar\intbl\tx630\adjustright ';

	    $ftr = '{\lang1024
{\shp{\*\shpinst'.$tpos.'\shpfhdr0\shpbxcolumn\shpbypage\shpwr3\shpwrk0\shpfblwtxt0\shpz0\shplid1026{\sp{\sn shapeType}{\sv 202}}{\sp{\sn fFlipH}{\sv 0}}{\sp{\sn fFlipV}{\sv 0}}
{\sp{\sn lTxid}{\sv 65536}}{\sp{\sn dxTextLeft}{\sv 64008}}{\sp{\sn dyTextTop}{\sv 0}}{\sp{\sn dxTextRight}{\sv 18288}}{\sp{\sn txflTextFlow}{\sv 2}}{\sp{\sn fLine}{\sv 0}}{\shptxt \pard\plain \widctlpar\adjustright \f16\fs18\cgrid {\fs'.intval($fs*4/3).'\f40
Page printed {\field{\*\fldinst DATE \\\\@ "M/d/yy" \\\\* MERGEFORMAT }}{\fldrslt }
\par }}}{\shprslt{\*\do\dobxcolumn\dobypage\dodhgt8192\dptxbx{\dptxbxtext\pard\plain \widctlpar\adjustright \f16\fs18\cgrid {\fs12 Page
\par }}\dpx-576\dpy10800\dpxsize288\dpysize1008\dpfillfgcr255\dpfillfgcg255\dpfillfgcb255\dpfillbgcr255\dpfillbgcg255\dpfillbgcb255\dpfillpat1\dplinehollow}}}}';

	    $rownum = 1;
	    $section = 0;
	    $page = 0;
	    $entrymin = ($prt_opt['general']['duplex'] ? 2*$prpg : $prpg);
	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);
    
	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		if (strtolower(substr($row[$sortkey],0,1)) >=
			strtolower(substr($sect_tabs[$section+1],0,1))) {
			
		    if ($rownum != 1) {
			while ($rownum <= $entrymin) {
			    if ($rownum == $prpg+1) {
				$page++;
				echo $hdrdef;
		    		$qr = (($page % 2 == 1) || !$prt_opt['general']['duplex']) ? '\qr' : '\ql';
				echo '\pard '.$qr.'\widctlpar\intbl\adjustright {\fs28\cf8 '.$sect_tabs[$section].'\cell }{\row }';
			    }
			    /* Generate blank entries until next odd page # */
			    echo $newrow . $rightcol . '\cell' . $endrow;
			    echo $newrow . $rightcol . '\cell' . $endrow;
			    echo $newrow . $rightcol . '\cell' . $endrow;
			    echo $newrow . $rightcol . '\cell' . $endrow;
			    echo $newrow2 . "{\\b Home}$rightcol{\\b Fax}\\cell $endrow";
			    echo ($rownum % $prpg != 0) ? $newrow3 :
			          str_replace("brdrs", "brdrnone", $newrow3);
			    echo "{\\b Office}$rightcol{\\b Other}\\cell $endrow";
			    $rownum++;
			}
			$rownum = 1;
		    }
		    while (strtolower(substr($row[$sortkey],0,1)) >
			   strtolower(substr($sect_tabs[$section],
			    strlen($sect_tabs[$section])-1,1)))
			$section++;
		}

		if ($rownum == 1 || $rownum == $prpg+1) {
		    /* New page */

		    $page++;
		    echo $hdrdef;
		    		$qr = (($page % 2 == 1) || !$prt_opt['general']['duplex']) ? '\qr' : '\ql';
		    echo '\pard '.$qr.'\widctlpar\intbl\adjustright {\fs' . ($fs+5)*2 . '\cf8 '.$sect_tabs[$section].'\cell }{\row }';
		    if ($qr == '\qr' || !$prt_opt['general']['duplex']) echo $ftr;
		}

		$td_out = $newrow;

		$td_out .= '\f1\fs' . ($fs+2)*2 . ' {\b ' . $row['name'] . "}";

		if (!empty($row['jobtitle']))
		    $td_out .= ', {\i\fs'. $fs*2 . ' ' . $row['jobtitle'] . "} ";
		$td_out .= $rightcol;

		if (!empty($row['company']))
		    $td_out .= "$row[company]";
		$td_out .= '\cell ';

	        $td_out .= $endrow . $newrow;

		$td_out .= str_replace("\n", ", ", $row['homestreet']);
		$td_out .= $rightcol . str_replace("\n", ", ",
			   $row['businessstreet']);
		$td_out .= '\cell ';

	        $td_out .= $endrow . $newrow;

		$comma = empty($row['homecity']) ? '' : ', ';
		$td_out .= "$row[homecity]$comma$row[homestate] $row[homezip]";
		$td_out .= $rightcol;
		$comma = empty($row['businesscity']) ? '' : ', ';
		$td_out .= "$row[businesscity]$comma$row[businessstate] $row[businesszip]";
		$td_out .= '\cell ';

	        $td_out .= $endrow . $newrow;

		$td_out .= $row['email'] . $rightcol;
		$td_out2 = '';
		if ($row['birthday'] != '0000-00-00')
		    $td_out2 .= 'Born: ' . date_fmt($row['birthday'], 'dd-mmm-yy');
		if (!empty($row['sigother']))
		    $td_out2 .= " SO: $row[sigother]";
		if ($row['anniversary'] != '0000-00-00')
		    $td_out2 .= ' {\i (A: '. date_fmt($row['anniversary'], 'dd-mmm-yy') . ')}';

		$td_out .= ltrim ($td_out2);
		$td_out .= '\cell ';

	        $td_out .= $endrow . $newrow2;

		$td_out .= '{\b Home}\tab ' . $row['homephone'] . $rightcol;
		$td_out .= '{\b Fax}\tab ' . $row['fax'] . '\cell ';

	        $td_out .= $endrow;
		echo ($rownum % $prpg != 0) ? $newrow3 :
			   str_replace("brdrs", "brdrnone", $newrow3);

		$td_out .= '\tx630{\b Office}\tab ' . $row['workphone'] . $rightcol;
		$td_out .= '{\b Cell}\tab ' . $row['cellphone'] . '\cell ';

		$td_out .= $endrow;

		echo $td_out;
		
		if (++$rownum > $entrymin) {
		    $rownum = 1;
		}
	    }
	    if ($rownum != 1) {
		while ($rownum <= $entrymin) {
		    if ($rownum == $prpg+1) {
			$page++;
			echo $hdrdef;
	    		$qr = ($page % 2 == 1) ? '\qr' : '\ql';
			echo '\pard '.$qr.'\widctlpar\intbl\adjustright {\fs28\cf8 '.$sect_tabs[$section].'\cell }{\row }';
		    }
		    /* Generate blank entries until next odd page # */
		    echo $newrow . $rightcol . '\cell' . $endrow;
		    echo $newrow . $rightcol . '\cell' . $endrow;
		    echo $newrow . $rightcol . '\cell' . $endrow;
		    echo $newrow . $rightcol . '\cell' . $endrow;
		    echo $newrow2 . "{\\b Home}$rightcol{\\b Fax}\\cell $endrow";
		    echo ($rownum % $prpg != 0) ? $newrow3 :
			   str_replace("brdrs", "brdrnone", $newrow3);
		    echo  "{\\b Office}$rightcol{\\b Other}\\cell $endrow";
		    $rownum++;
		}
	    }
	    echo '\pard \widctlpar\intbl\adjustright {\v }}';
	    break;

	/*
	 * CSV
	 */
	case "print $printformats[11]":
	    $fields = 'uid,First Name,Middle Name,Last Name,Company,Job Title,Home Street,Home City,Home State,Home Zip,Home Country,Business Street,Business City,Business State,Business Zip,Business Country,Other Street,Other City,Other State,Other Zip,Other Country,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Work Phone,Cell Phone,Fax,Address Sel,Email 1,Email 2,Email 3,Email Sel,Sig Other,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';
	    $ms_names = 'uid,First Name,Middle Name,Last Name,Company,Job Title,Home Street,Home City,Home State,Home Postal Code,Home Country,Business Street,Business City,Business State,Business Postal Code,Business Country,Other Street,Other City,Other State,Other Postal Code,Other Country,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Business Phone,Mobile Phone,Business Fax,Address Sel,E-mail Address,E-mail 2 Address,E-mail 3 Address,Email Sel,Spouse,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';
	    echo "$ms_names\n";
	    $f = explode(',', $fields);

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);
		$td_out = '';
		$comma = '';
		foreach ($f as $field) {
		    $td_out .= $comma;
		    $field = strtolower(str_replace(' ','',$field));
		    if (!empty ($row[$field])) {
			if ($field == 'categories')
			    $td_out .= '"' . str_replace (',', ';', $row['categories']) . '"';
			else if ($field == 'modified' || $field == 'created')
			    $td_out .= '"' . date_fmt_sql ($row[$field], 'Y-m-d h:i:s') . '"';
			else
			    $td_out .= '"' . str_replace ('"', '""', $row[$field]) . '"';
		    }
		    $comma = ',';
		}
		echo $td_out . "\n";
	    }
	    break;

	/*
	 * Vcards
	 */
	case "print $printformats[12]":
	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);

		$td_out = "BEGIN:VCARD \r\nVERSION:2.1\r\n";
		$td_out .= "UID:$row[uid]\r\n";
		$td_out .= "N:$row[lastname];$row[firstname]\r\n";
		$td_out .= "FN:$row[name]\r\n";
		for ($i = 1; $i <= 5; $i++) {
		    if (!empty($row['phone'.$i])) {
			$td_out .= 'TEL;TYPE=';
			if ($row['phonesel'] == $i)
			    $td_out .= 'PREF,';
			$td_out .= strtoupper($phonelabels[$row['phone'.$i.'type']]);
			$p = $row['phone'.$i];
			if (substr($p,0,1) == '(' && substr($p,4,1) == ')')
			    $td_out .= ':+1-'.substr($p,1,3).'-'.
				trim(substr($p,5)) . "\r\n";
			else
			    $td_out .= ":" . $row['phone'.$i] . "\r\n";
		    }
		}
		if (!empty($row['webpage']))
		    $td_out .= "URL:$row[webpage]\r\n";
		if (!empty($row['company']))
		    $td_out .= "ORG:$row[company];\r\n";
		if (!empty($row['jobtitle']))
		    $td_out .= "TITLE:$row[jobtitle]\r\n";
		if (!empty($row['notes']))
		    $td_out .= "NOTE".need_qp_enc($row['notes']).":".
				qp_enc($row['notes'])."\r\n";
		for ($i = 1; $i <=3; $i++) {
		    if (!empty($row['email'.$i])) {
		        $td_out .= "EMAIL;TYPE=";
			if ($row['emailsel'] == $i)
			    $td_out .= "PREF,";
			$td_out .= "INTERNET:".$row['email'.$i]."\r\n";
		    }
		}
		if (!empty($row['businessaddress'])) {
		    $addr = str_replace(";", ",", $row['businessstreet']).
			";$row[businesscity];$row[businessstate];$row[businesszip];$row[businesscountry]";
		    $td_out .= "ADR;TYPE=WORK".need_qp_enc($addr).":;;".qp_enc($addr)."\r\n";
		}
		if (!empty($row['homeaddress'])) {
		    $addr = str_replace(";", ",", $row['homestreet']).
			";$row[homecity];$row[homestate];$row[homezip];$row[homecountry]";
		    $pref = ($row['addresssel'] == 'H') ? ',PREF' : '';
		    $td_out .= "ADR;TYPE=HOME".$pref.need_qp_enc($addr).":;;".qp_enc($addr)."\r\n";
		}
		if (!empty($row['otheraddress'])) {
		    $addr = str_replace(";", ",", $row['otherstreet']).
			";$row[othercity];$row[otherstate];$row[otherzip];$row[othercountry]";
		    $td_out .= "ADR;TYPE=POSTAL".need_qp_enc($addr).":;;".qp_enc($addr)."\r\n";
		}
		if ($row['birthday'] != '0000-00-00')
		    $td_out .= "BDAY:$row[birthday]\r\n";

		$td_out .= "REV:".substr($row['modified'],0,8).'T'.
			   substr($row['modified'],8,6)."\r\n";
		if (!empty($row['categories']))
		    $td_out .= "CATEGORIES:$row[categories]\r\n";

		/*
		 * Some of our fields are non-standard.
		 */
		if (!empty($row['sigother']))
		    $td_out .= "X-SPOUSE:$row[sigother]\r\n";
		if ($row['anniversary'] != '0000-00-00')
		    $td_out .= "X-ANNIV:$row[anniversary]\r\n";
		if (!empty($row['children']))
		    $td_out .= "X-CHILDREN:$row[children]\r\n";

		$td_out .= "END:VCARD\r\n\r\n";
		echo $td_out;
	    }
	    break;

	/*
	 * Birthday & anniversary list
	 */
	case "print $printformats[13]":
	    $dates = array();

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);
		if ($row['birthday'] != '0000-00-00')
		    $dates[substr($row['birthday'],5,5)] .= "$row[name]:".
			substr($row['birthday'],0,4).":B;";
		if ($row['anniversary'] != '0000-00-00')
		    $dates[substr($row['anniversary'],5,5)] .= "$row[name] & ".
		    "$row[sigother]:" . substr($row['anniversary'],0,4).":A;";
	    }
	    echo rtf_header($prt_opt['general']['form_factor'] == 'full');
	    echo rtf_page_format($prt_opt['general']['form_factor'], $prt_opt);
	    echo '\fs20';
	    ksort ($dates);
	    $month = 0;
	    foreach ( $dates as $day => $entry ) {
		if (substr($day,0,2) != $month) {
		    echo '\pard \keepn\widctlpar\adjustright '.
		     '\shading2500\cbpat8{\b\fs24 '.$monthnames[$month]."\n" .
		     '\par }\pard \widctlpar\fi-480\li480\adjustright ';
		    $month = ltrim(substr($day,0,2), '0');
		}
		$items = explode(";", substr($entry,0,strlen($entry)-1));
		echo '{\b ' . ltrim(substr($day,3,2), '0') . '}\tab ';
		$sep = '';
		foreach ( $items as $item) {
		    list ($name, $year, $event) = explode(':', $item);
		    echo "$sep$name";
		    if ($year != NULL_YEAR)
			echo " {\fs16\i (".($event == 'A' ? 'together ' : '').
			     "$year)}";
		    $sep = "\n\\par \\tab ";
		}
		echo "\n\\par ";
	    }
	    echo "}}";
	    break;

	/*
	 * Nokia CSV
	 */
	case "print $printformats[14]":
	    $fields = 'uid,First Name,Middle Name,Last Name,Company,Job Title,Home Street,Home City,Home State,Home Zip,Home Country,Business Street,Business City,Business State,Business Zip,Business Country,Other Street,Other City,Other State,Other Zip,Other Country,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Work Phone,Cell Phone,Fax,Address Sel,Email 1,Email 2,Email 3,Email Sel,Sig Other,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';
	    $ms_names = 'uid,First Name,Middle Name,Last Name,Company,Job Title,Home Street,Home City,Home State,Home Postal Code,Home Country,Business Street,Business City,Business State,Business Postal Code,Business Country,Other Street,Other City,Other State,Other Postal Code,Other Country,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Business Phone,Mobile Phone,Business Fax,Address Sel,E-mail Address,E-mail 2 Address,E-mail 3 Address,Email Sel,Spouse,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';
//	    $fields = 'uid,First Name,Middle Name,Last Name,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Work Phone,Cell Phone,Fax,Email 1,Email 2,Email 3,Email Sel,Birthday,Categories';
//	    $ms_names = 'uid,First Name,Middle Name,Last Name,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Business Phone,Mobile Phone,Business Fax,E-mail Address,E-mail 2 Address,E-mail 3 Address,Email Sel,,Birthday,Categories';
	    echo "$ms_names\n";
	    $f = explode(',', $fields);

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);
		$td_out = '';
		$comma = '';
		foreach ($f as $field) {
		    $td_out .= $comma;
		    $field = strtolower(str_replace(' ','',$field));
		    if (!empty ($row[$field])) {
			if ($field == 'categories')
			    $td_out .= '"' . str_replace (',', ';', $row[categories]) . '"';
			else if ($field == 'modified' || $field == 'created')
			    $td_out .= '"' . date_fmt_sql ($row[$field], 'Y-m-d h:i:s') . '"';
			else if ($field == 'lastname')
			    $td_out .= '"' . "$row[lastname]/$row[firstname]" . '"';
			else if ($field == 'firstname')
			    $td_out .= '';
			else
			    $td_out .= '"' . str_replace ('"', '""', $row[$field]) . '"';
		    }
		    $comma = ',';
		}
		echo $td_out . "\n";
	    }
	    break;

	/*
	 * Motorola CSV
	 */
	case "print $printformats[15]":
	    $fields = 'uid,First Name,Middle Name,Last Name,Company,Job	Title,Home Street,Home City,Home State,Home Zip,Home Country,Business Street,Business City,Business State,Business Zip,Business Country,Home Phone,Work Phone,Cell Phone,Fax,Phone 5,Email 1,Email 2,Sig Other,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';
	    $mot_names = 'PrivateID,FirstName,MiddleName,LastName,Company,JobTitle,Address_D,City_D,State_D,ZCode_D,Country_D,Address_B,City_B,State_B,ZCode_B,Country_B,Tel_D,Tel.RD,Cell_D,Fax_B,Other,Email_D,Email_B,Spouse,Children,Anniversary,Birthday,WebPage,User.1,User.2,User.3,User.4,Categories,Notes,Modified,Created';
	    echo "$mot_names\n";
	    $f = explode(',', $fields);

	    foreach ( $sel as $entry ) {
		list($sbackend, $uid) = explode(':', $entry);

	        /* Retrieve all fields of the current record */

		$row = $abook->lookup($uid, $sbackend);
		$td_out = '';
		$comma = '';
		foreach ($f as $field) {
		    $td_out .= $comma;
		    $field = strtolower(str_replace(' ','',$field));

		    /* Strip newlines and commas to avoid confusing parser */
		    $row[$field] = str_replace("\r", ' ', $row[$field]);
		    $row[$field] = str_replace("\n", ' ', $row[$field]);
		    $row[$field] = str_replace(",", ';', $row[$field]);
		    if (!empty ($row[$field])) {
			if ($field == 'categories')
			    $td_out .= '"' . str_replace (',', ';', $row[categories]) . '"';
			else if ($field == 'modified' || $field == 'created')
			    $td_out .= '"' . date_fmt_sql ($row[$field], 'Y-m-d h:i:s') . '"';
			else
			    $td_out .= '"' . str_replace ('"', '""', $row[$field]) . '"';
		    }
		    $comma = ',';
		}
		echo $td_out . "\r\n";
	    }
	    break;

	default:
	    $formerror = "Please choose a valid print format ($prtfmt)<br>";
	    echo html_tag( 'table',
	        html_tag( 'tr',
	            html_tag( 'td',
	                   "\n". '<br><strong><font color="' . $color[2] .
	                   '">' . _("ERROR") . ': ' . $formerror . '</font></strong>' ."\n",
	            'center' )
	        ),
	    'center', '', 'width="100%"' );
    }
    break;
}

    if ($showform != 'print') {
	/* Add hook for anything that wants on the bottom */
	do_hook('addressbook_bottom');

	echo '</BODY></HTML>';
    }
?>
