<?php

/**
 * options/addrbook.php
 * $Id: addrbook.php,v 1.6 2011/10/14 15:56:54 richb Exp $
 *
 * Written by Rich Braun 18 Jun 2003
 */

/* SquirrelMail required files. */
/*  (none) */

/* Define the group constants for the personal options page. */
define('SMOPT_GRP_ADDRDISP',  0);
define('SMOPT_GRP_ADDRPRINT', 1);
define('SMOPT_GRP_ADDRENV',   2);
define('SMOPT_GRP_ADDRLBL',   3);
define('SMOPT_GRP_ADDRSORT',  4);

$addr_fields = 'uid,First Name,Middle Name,Last Name,Company,Job Title,Home Street,Home City,Home State,Home Zip,Home Country,Business Street,Business City,Business State,Business Zip,Business Country,Other Street,Other City,Other State,Other Zip,Other Country,Phone 1,Phone 2,Phone 3,Phone 4,Phone 5,Phone 1 Type,Phone 2 Type,Phone 3 Type,Phone 4 Type,Phone 5 Type,Phone Sel,Home Phone,Work Phone,Cell Phone,Fax,Address Sel,Email 1,Email 2,Email 3,Email Sel,Sig Other,Children,Anniversary,Birthday,Web Page,User 1,User 2,User 3,User 4,Categories,Notes,Modified,Created';

/* Special explode function to set key,val in array, squeezing out spaces */
function explode_key($sep, $string) {
   $tmp = explode($sep, $string);
   $ret = array();
   $ret[' '] = '';   // Special case:  distinguish blank from un-set
   foreach ($tmp as $val)
       $ret[strtolower(str_replace(' ','',$val))] = $val;
   return $ret;
}

/* Define the optpage load function for the personal options page. */
function load_optpage_data_addrbook() {
    global $data_dir, $username, $addr_fields,
	$addr_list_numrows, $addr_custom_labels, 
        $addr_list_columns,
	$mailing_address,
	$addr_hide_other,
	$addr_lastname_first,
	$addr_highlight_cat,
	$addr_sort_1_field,
	$addr_sort_1_dir,
	$addr_sort_2_field,
	$addr_sort_2_dir,
	$addr_sort_3_field,
	$addr_sort_3_dir,
	$addr_sort_4_field,
	$addr_sort_4_dir,
	$addr_sort_5_field,
	$addr_sort_5_dir,
	$rpt_gen_sigother,
	$rpt_gen_suppress_comma,
	$rpt_gen_duplex,
	$rpt_gen_form_factor,
	$rpt_wpages_sectmin,
	$rpt_detail_sectmin,
	$rpt_env_returnaddr,
	$rpt_env_barcode,
	$rpt_env_fontname,
	$rpt_env_fontsize,
	$rpt_lbl_barcode,
	$rpt_lbl_fontname,
	$rpt_lbl_fontsize;

    $addr_list_columns = getPref($data_dir, $username, 'addr_list_columns');
    $addr_list_numrows = getPref($data_dir, $username, 'addr_list_numrows');
    $addr_custom_labels = getPref($data_dir, $username, 'addr_custom_labels');
    $addr_hide_other = getPref($data_dir, $username, 'addr_hide_other');
    $addr_lastname_first = getPref($data_dir, $username, 'addr_lastname_first');
    $addr_highlight_cat = getPref($data_dir, $username, 'addr_highlight_cat');
    $addr_sort_1_field = getPref($data_dir, $username, 'addr_sort_1_field');
    $addr_sort_1_dir = getPref($data_dir, $username, 'addr_sort_1_dir');
    $addr_sort_2_field = getPref($data_dir, $username, 'addr_sort_2_field');
    $addr_sort_2_dir = getPref($data_dir, $username, 'addr_sort_2_dir');
    $addr_sort_3_field = getPref($data_dir, $username, 'addr_sort_3_field');
    $addr_sort_3_dir = getPref($data_dir, $username, 'addr_sort_3_dir');
    $addr_sort_4_field = getPref($data_dir, $username, 'addr_sort_4_field');
    $addr_sort_4_dir = getPref($data_dir, $username, 'addr_sort_4_dir');
    $addr_sort_5_field = getPref($data_dir, $username, 'addr_sort_5_field');
    $addr_sort_5_dir = getPref($data_dir, $username, 'addr_sort_5_dir');
    $mailing_address = getPref($data_dir, $username, 'mailing_address');
    $rpt_gen_sigother = getPref($data_dir, $username, 'rpt_gen_sigother');
    $rpt_gen_suppress_comma = getPref($data_dir, $username, 'rpt_gen_suppress_comma');
    $rpt_gen_duplex = getPref($data_dir, $username, 'rpt_gen_duplex');
    $rpt_gen_form_factor = getPref($data_dir, $username, 'rpt_gen_form_factor');
    $rpt_wpages_sectmin = getPref($data_dir, $username, 'rpt_wpages_sectmin');
    $rpt_detail_sectmin = getPref($data_dir, $username, 'rpt_detail_sectmin');
    $rpt_env_returnaddr = getPref($data_dir, $username, 'rpt_env_returnaddr');
    $rpt_env_barcode = getPref($data_dir, $username, 'rpt_env_barcode');
    $rpt_env_fontname = getPref($data_dir, $username, 'rpt_env_fontname');
    $rpt_env_fontsize = getPref($data_dir, $username, 'rpt_env_fontsize');
    $rpt_lbl_barcode = getPref($data_dir, $username, 'rpt_lbl_barcode');
    $rpt_lbl_fontname = getPref($data_dir, $username, 'rpt_lbl_fontname');
    $rpt_lbl_fontsize = getPref($data_dir, $username, 'rpt_lbl_fontsize');

    /* Defaults */
    if (empty ($addr_list_columns))
	$addr_list_columns = "Name,Email,Company,Phone";
    if (empty ($addr_custom_labels))
	$addr_custom_labels = "'Custom 1','Custom 2','Custom 3','Custom 4'";
    if (empty ($addr_sort_1_field))
	$addr_sort_1_field = "lastname";
    if (empty ($addr_sort_2_field))
	$addr_sort_2_field = "firstname";
    if (empty ($rpt_env_fontsize))
	$rpt_env_fontsize = 12;
    if (empty ($rpt_lbl_fontsize))
	$rpt_lbl_fontsize = 10;

    /* Build a simple array into which we will build options. */
    $optgrps = array();
    $optvals = array();

    /******************************************************/
    /* LOAD EACH GROUP OF OPTIONS INTO THE OPTIONS ARRAY. */
    /******************************************************/

    /*** Load the Contact Information Options into the array ***/
    $optgrps[SMOPT_GRP_ADDRDISP] = _("Address List Display Options");
    $optvals[SMOPT_GRP_ADDRDISP] = array();

    /* Build a simple array into which we will build options. */
    $optvals = array();

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_list_columns',
        'caption' => _("Columns to Display"),
        'type'    => SMOPT_TYPE_INTEGER,
        'refresh' => SMOPT_REFRESH_ALL,
        'size'    => SMOPT_SIZE_HUGE
    );

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_list_numrows',
        'caption' => _("Number of Display Rows"),
        'type'    => SMOPT_TYPE_INTEGER,
        'refresh' => SMOPT_REFRESH_ALL,
        'size'    => SMOPT_SIZE_SMALL
    );

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_hide_other',
        'caption' => _("Hide <i>other address</i> on entry form"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL,
    );

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_lastname_first',
        'caption' => _("Display last name first"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL,
    );

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_highlight_cat',
	'caption' => _("Highlight names in category"),
	'type'	  => SMOPT_TYPE_STRING,
	'refresh' => SMOPT_REFRESH_ALL,
    );

    $optvals[SMOPT_GRP_ADDRDISP][] = array(
        'name'    => 'addr_custom_labels',
        'caption' => _("Custom label names"),
        'type'    => SMOPT_TYPE_STRING,
        'refresh' => SMOPT_REFRESH_ALL,
        'size'    => SMOPT_SIZE_HUGE
    );

    /*** Load the Printing Options into the array ***/
    $optgrps[SMOPT_GRP_ADDRPRINT] = _("Printing Options");
    $optvals[SMOPT_GRP_ADDRPRINT] = array();

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'mailing_address',
        'caption' => _("Return address"),
        'type'    => SMOPT_TYPE_TEXTAREA,
        'refresh' => SMOPT_REFRESH_ALL,
        'size'    => SMOPT_SIZE_SMALL
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_gen_sigother',
        'caption' => _("Combine SigOther"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_gen_suppress_comma',
        'caption' => _("Suppress Comma Between City/State"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_gen_duplex',
        'caption' => _("Duplex (2-sided) Printing"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_gen_form_factor',
        'caption' => _("Form Factor"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('full'         => _("Full 8.5 x 11"),
                           'half'         => _("DayRunner 5.5 x 8.5"),
                           'pers'         => _("DayRunner 3.75 x 6.75"))
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_wpages_sectmin',
        'caption' => _("Section Headers for White Pages--Min Records"),
        'type'    => SMOPT_TYPE_INTEGER,
        'size'    => SMOPT_SIZE_SMALL,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRPRINT][] = array(
        'name'    => 'rpt_detail_sectmin',
        'caption' => _("Section Headers for Detail Report--Min Records"),
        'type'    => SMOPT_TYPE_INTEGER,
        'size'    => SMOPT_SIZE_SMALL,
        'refresh' => SMOPT_REFRESH_ALL
    );

    /*** Load the Envelope Options into the array ***/
    $optgrps[SMOPT_GRP_ADDRENV] = _("Envelope Options");
    $optvals[SMOPT_GRP_ADDRENV] = array();

    $optvals[SMOPT_GRP_ADDRENV][] = array(
        'name'    => 'rpt_env_returnaddr',
        'caption' => _("Print Return Address"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRENV][] = array(
        'name'    => 'rpt_env_barcode',
        'caption' => _("Barcode"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRENV][] = array(
        'name'    => 'rpt_env_fontname',
        'caption' => _("Font Name"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('Arial'         => _("Arial"),
                           'Arial Narrow'  => _("Arial Narrow"),
                           'Courier New'   => _("Courier New"),
                           'Comic Sans MS' => _("Comic Sans MS"),
                           'Times New Roman' => _("Times New Roman") )
    );

    $optvals[SMOPT_GRP_ADDRENV][] = array(
        'name'    => 'rpt_env_fontsize',
        'caption' => _("Font Size"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('8'         => _("8"),
                           '8.5'       => _("8.5"),
                           '9'         => _("9"),
                           '9.5'       => _("9.5"),
                           '10'        => _("10"),
                           '10.5'      => _("10.5"),
                           '11'        => _("11"),
                           '11.5'      => _("11.5"),
                           '12'        => _("12"),
                           '12.5'      => _("12.5"),
                           '13'        => _("13"),
                           '14'        => _("14"),
                           '15'        => _("15") )
    );

    /*** Load the Mailing Labels Options into the array ***/
    $optgrps[SMOPT_GRP_ADDRLBL] = _("Mailing Labels Options");
    $optvals[SMOPT_GRP_ADDRLBL] = array();

    $optvals[SMOPT_GRP_ADDRLBL][] = array(
        'name'    => 'rpt_lbl_barcode',
        'caption' => _("Barcode"),
        'type'    => SMOPT_TYPE_BOOLEAN,
        'refresh' => SMOPT_REFRESH_ALL
    );

    $optvals[SMOPT_GRP_ADDRLBL][] = array(
        'name'    => 'rpt_lbl_fontname',
        'caption' => _("Font Name"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('Arial'         => _("Arial"),
                           'Arial Narrow'  => _("Arial Narrow"),
                           'Courier New'   => _("Courier New"),
                           'Comic Sans MS' => _("Comic Sans MS"),
                           'Times New Roman' => _("Times New Roman") )
    );

    $optvals[SMOPT_GRP_ADDRLBL][] = array(
        'name'    => 'rpt_lbl_fontsize',
        'caption' => _("Font Size"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('8'         => _("8"),
                           '8.5'       => _("8.5"),
                           '9'         => _("9"),
                           '9.5'       => _("9.5"),
                           '10'        => _("10"),
                           '10.5'      => _("10.5"),
                           '11'        => _("11"),
                           '11.5'      => _("11.5"),
                           '12'        => _("12"),
                           '12.5'      => _("12.5"),
                           '13'        => _("13") )
    );

    /*** Load the Sort Options into the array ***/
    $optgrps[SMOPT_GRP_ADDRSORT] = _("Sort Options");
    $optvals[SMOPT_GRP_ADDRSORT] = array();

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_1_field',
        'caption' => _("Sort By"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => explode_key(',', $addr_fields)
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_1_dir',
        'caption' => _("Direction"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('ASC'    => _("Ascending"),
                           'DESC'   => _("Descending"))
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_2_field',
        'caption' => _("Then By"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => explode_key(',', $addr_fields)
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_2_dir',
        'caption' => _("Direction"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('ASC'    => _("Ascending"),
                           'DESC'   => _("Descending"))
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_3_field',
        'caption' => _("Then By"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => explode_key(',', $addr_fields)
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_3_dir',
        'caption' => _("Direction"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('ASC'    => _("Ascending"),
                           'DESC'   => _("Descending"))
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_4_field',
        'caption' => _("Then By"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => explode_key(',', $addr_fields)
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_4_dir',
        'caption' => _("Direction"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('ASC'    => _("Ascending"),
                           'DESC'   => _("Descending"))
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_5_field',
        'caption' => _("Then By"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => explode_key(',', $addr_fields)
    );

    $optvals[SMOPT_GRP_ADDRSORT][] = array(
        'name'    => 'addr_sort_5_dir',
        'caption' => _("Direction"),
        'type'    => SMOPT_TYPE_STRLIST,
        'refresh' => SMOPT_REFRESH_ALL,
        'posvals' => array('ASC'    => _("Ascending"),
                           'DESC'   => _("Descending"))
    );

    /* Assemble all this together and return it as our result. */
    $result = array(
        'grps' => $optgrps,
        'vals' => $optvals
    );
    return ($result);
}

/******************************************************************/
/** Define any specialized save functions for this option page. ***/
/******************************************************************/

?>
