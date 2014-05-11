<?php
   
/**
 * abook_database.php
 *
 * Copyright (c) 1999-2003 The SquirrelMail Project Team
 * Licensed under the GNU GPL. For full terms see the file COPYING.
 *
 * Backend for personal addressbook stored in a database,
 * accessed using the DB-classes in PEAR.
 *
 * IMPORTANT:  The PEAR modules must be in the include path
 * for this class to work.
 *
 * An array with the following elements must be passed to
 * the class constructor (elements marked ? are optional):
 *
 *    dsn       => database DNS (see PEAR for syntax)
 *    table     => table to store addresses in (must exist)
 *    owner     => current user (owner of address data)
 *  ? writeable => set writeable flag (true/false)
 *
 * The table used should have the following columns:
 * owner, nickname, firstname, lastname, email, label
 * The pair (owner,nickname) should be unique (primary key).
 *
 *  NOTE. This class should not be used directly. Use the
 *        "AddressBook" class instead.
 *
 * $Id: abook_database.php,v 1.27 2011/10/21 18:12:50 richb Exp $
 */
   
require_once('DB.php');

/*
 * Create the tables in MySQL table thus:
 *
CREATE TABLE contacts_username (
  uid varchar(10) NOT NULL default '',
  created_by varchar(16) NOT NULL default 'richb',
  Modified timestamp,
  Created timestamp,
  FirstName varchar(32) NOT NULL default '',
  MiddleName varchar(32) NOT NULL default '',
  LastName varchar(32) NOT NULL default '',
  Title varchar(32) NOT NULL default '',
  Suffix varchar(32) NOT NULL default '',
  NickName varchar(32) NOT NULL default '',
  Gender char(1) NOT NULL default '',
  Company varchar(64) NOT NULL default '',
  JobTitle varchar(32) NOT NULL default '',
  Department varchar(64) NOT NULL default '',
  BusinessStreet varchar(128) NOT NULL default '',
  BusinessCity varchar(32) NOT NULL default '',
  BusinessState varchar(16) NOT NULL default '',
  BusinessPostalCode varchar(10) NOT NULL default '',
  BusinessCountry varchar(32) NOT NULL default '',
  WebPage varchar(64) NOT NULL default '',
  HomeStreet varchar(128) NOT NULL default '',
  HomeCity varchar(32) NOT NULL default '',
  HomeState varchar(16) NOT NULL default '',
  HomePostalCode varchar(10) NOT NULL default '',
  HomeCountry varchar(32) NOT NULL default '',
  OtherStreet varchar(128) NOT NULL default '',
  OtherCity varchar(32) NOT NULL default '',
  OtherState varchar(16) NOT NULL default '',
  OtherPostalCode varchar(10) NOT NULL default '',
  OtherCountry varchar(32) NOT NULL default '',
  AddressSelector char(1) NOT NULL default 'H',
  Phone1 varchar(20) NOT NULL default '',
  Phone1Type int(1) NOT NULL default '0',
  Phone2 varchar(20) NOT NULL default '',
  Phone2Type int(1) NOT NULL default '1',
  Phone3 varchar(20) NOT NULL default '',
  Phone3Type int(1) NOT NULL default '2',
  Phone4 varchar(20) NOT NULL default '',
  Phone4Type int(1) NOT NULL default '7',
  Phone5 varchar(20) NOT NULL default '',
  Phone5Type int(1) NOT NULL default '3',
  PhoneDisplay int(1) NOT NULL default '2',
  AssistantsName varchar(64) NOT NULL default '',
  AssistantsPhone varchar(20) NOT NULL default '',
  ManagersName varchar(64) NOT NULL default '',
  CompanyMainPhone varchar(20) NOT NULL default '',
  OfficeLocation varchar(64) NOT NULL default '',
  Account varchar(32) NOT NULL default '',
  Email1 varchar(64) NOT NULL default '',
  Email2 varchar(64) NOT NULL default '',
  Email3 varchar(64) NOT NULL default '',
  EmailSelector int(1) NOT NULL default '1',
  Anniversary date,
  Birthday date,
  SigOther varchar(64) NOT NULL default '',
  Children varchar(64) NOT NULL default '',
  Hobby varchar(64) NOT NULL default '',
  Categories set('Personal', 'Professional') NOT NULL default '',
  User1 varchar(64) NOT NULL default '',
  User2 varchar(64) NOT NULL default '',
  User3 varchar(64) NOT NULL default '',
  User4 varchar(64) NOT NULL default '',
  Photo mediumblob NOT NULL,
  Notes mediumtext,
  PRIMARY KEY (uid)
);

CREATE TABLE history_username (
  uid varchar(10) NOT NULL default '',
  moddate timestamp(14) NOT NULL,
  modby varchar(16) NOT NULL default '',
  field varchar(20) NOT NULL default '',
  newval varchar(255) NOT NULL default '',
  oldval varchar(255) NOT NULL default ''
);

*/

/* Phone label codes--0 thru 7 except 4 are defined by PalmOS */

define('PHONE_LABEL_WORK',   0);
define('PHONE_LABEL_HOME',   1);
define('PHONE_LABEL_FAX',    2);
define('PHONE_LABEL_OTHER',  3);
define('PHONE_LABEL_MAIN',   5);
define('PHONE_LABEL_PAGER',  6);
define('PHONE_LABEL_MOBILE', 7);
define('PHONE_LABEL_WORK2',  8);
define('PHONE_LABEL_HOME2',  9);


class abook_database extends addressbook_backend {
    var $btype = 'local';
    var $bname = 'database';
      
    var $dsn       = '';
    var $table     = '';
    var $history   = '';
    var $owner     = '';
    var $dbh       = false;

    var $writeable = true;
      
    var $phonelabelnames = array(
		PHONE_LABEL_WORK   => 'W',
		PHONE_LABEL_HOME   => 'H',
		PHONE_LABEL_FAX    => 'F',
		PHONE_LABEL_OTHER  => 'O',
		PHONE_LABEL_MAIN   => 'M',
		PHONE_LABEL_PAGER  => 'P',
		PHONE_LABEL_MOBILE => 'C',
		PHONE_LABEL_WORK2  => 'W2',
		PHONE_LABEL_HOME2  => 'H2');

    /*
     * Field names:
     * These fields are written to the MySQL contacts table.
     * They are customized for this application by the author but more or less
     * match the names used in such applications as Outlook and Palm Desktop.
     * One additional field is written by this module, Created, which is a
     * timestamp generated by the add() function.
     */

    /*           Field         MySQL name */
    var $fnames = array( 'uid'      => 'uid',
		'firstname' => 'FirstName',
		'lastname'  => 'LastName',
		'company'   => 'Company',
		'jobtitle'  => 'JobTitle',
		'homestreet'=> 'HomeStreet',
		'homecity'  => 'HomeCity',
		'homestate' => 'HomeState',
		'homezip'   => 'HomePostalCode',
		'homecountry' => 'HomeCountry',
		'businessstreet'=> 'BusinessStreet',
		'businesscity'  => 'BusinessCity',
		'businessstate' => 'BusinessState',
		'businesszip'   => 'BusinessPostalCode',
		'businesscountry' => 'BusinessCountry',
		'otherstreet'=> 'OtherStreet',
		'othercity'  => 'OtherCity',
		'otherstate' => 'OtherState',
		'otherzip'   => 'OtherPostalCode',
		'othercountry' => 'OtherCountry',
		'phone1'    => 'Phone1',
		'phone2'    => 'Phone2',
		'phone3'    => 'Phone3',
		'phone4'    => 'Phone4',
		'phone5'    => 'Phone5',
		'phone1type'  => 'Phone1Type',
		'phone2type'  => 'Phone2Type',
		'phone3type'  => 'Phone3Type',
		'phone4type'  => 'Phone4Type',
		'phone5type'  => 'Phone5Type',
		'phonesel'  => 'PhoneDisplay',
		'addresssel' => 'AddressSelector',
		'email1'    => 'Email1',
		'email2'    => 'Email2',
		'email3'    => 'Email3',
		'emailsel'  => 'EmailSelector',
		'sigother'  => 'SigOther',
		'children'  => 'Children',
		'anniversary' => 'Anniversary',
		'birthday'  => 'Birthday',
		'webpage'   => 'WebPage',
		'user1'     => 'User1',
		'user2'     => 'User2',
		'user3'     => 'User3',
		'user4'     => 'User4',
		'categories' => 'Categories',
		'notes'     => 'Notes');

    /* ========================== Private ======================= */
      
    /* Constructor */
    function abook_database($param) {
        $this->sname = _("Personal address book");
         
        if (is_array($param)) {
            if (empty($param['dsn']) || 
                empty($param['table']) || 
                empty($param['owner']) ||
		empty($param['history'])) {
                return $this->set_error('Invalid parameters');
            }
            
            $this->dsn   = $param['dsn'];
            $this->table = $param['table'];
            $this->owner = $param['owner'];
            $this->history = $param['history'];
	    if (!empty($param['sortby']))
		$this->sortby = $param['sortby'];
            
            if (!empty($param['name'])) {
               $this->sname = $param['name'];
            }

            if (isset($param['writeable'])) {
               $this->writeable = $param['writeable'];
            }

            $this->open(true);
        }
        else {
            return $this->set_error('Invalid argument to constructor');
        }
    }
      
      
    /* Open the database. New connection if $new is true */
    function open($new = false) {
        $this->error = '';
         
        /* Return true is file is open and $new is unset */
        if ($this->dbh && !$new) {
            return true;
        }
         
        /* Close old file, if any */
        if ($this->dbh) {
            $this->close();
        }
         
        $dbh = DB::connect($this->dsn, true);
         
        if (DB::isError($dbh)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($dbh)));
        }
         
        $this->dbh = $dbh;
        return true;
    }

    /* Close the file and forget the filehandle */
    function close() {
        $this->dbh->disconnect();
        $this->dbh = false;
    }

    /* ========================== Public ======================== */
     
    /* Search the file */
    function &search($expr) {
        $ret = array();
        if(!$this->open()) {
            return false;
        }
         
        /* To be replaced by advanded search expression parsing */
        if (is_array($expr)) {
            return;
        }

        /* Make regexp from glob'ed expression  */
        $expr = str_replace('?', '_', $expr);
        $expr = str_replace('*', '%', $expr);
        $expr = $this->dbh->quoteString($expr);
        $expr = "%$expr%";

        $query = sprintf("SELECT * FROM %s WHERE created_by='%s' AND " .
                         "(FirstName LIKE '%s' OR LastName LIKE '%s')" .
			 " ORDER BY %s",
                         $this->table, $this->owner, $expr, $expr, $this->sortby);
        $res = $this->dbh->query($query);

        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }

        while ($row = $res->fetchRow(DB_FETCHMODE_ASSOC)) {
	    $phone = $row['Phone'.$row['PhoneDisplay']];
	    $phonelabel = $row['Phone'.$row['PhoneDisplay'].'Type'];
	    if ( !empty( $phone ) )
	      $phone = $phone . "(".$this->phonelabelnames[$phonelabel].")";
            if ( !empty( $row[EmailSelector] ) )
	      $email = $row['Email'.$row[EmailSelector]];
	    else
	      $email = $row[Email1];

            array_push($ret, array('uid'  => $row[uid],
                                   'name'      => $row['FirstName']." ".$row['LastName'],
                                   'firstname' => $row['FirstName'],
                                   'lastname'  => $row['LastName'],
                                   'email'     => $email,
                                   'phonedisplay' => $phone,
                                   'backend'   => $this->bnum,
                                   'source'    => &$this->sname));
        }
        return $ret;
    }
     
    /* Lookup alias */
    function &lookup($alias) {

        if (empty($alias)) {
            return array();
        }
         
        $alias = strtolower($alias);

        if (!$this->open()) {
            return false;
        }
         
        $query = sprintf("SELECT * FROM %s WHERE created_by='%s' AND uid='%s'",
                         $this->table, $this->owner, $alias);

        $res = $this->dbh->query($query);

        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }

        if ($row = $res->fetchRow(DB_FETCHMODE_ASSOC)) {

	    /* Construct composite fields:  primary phone, primary email,
	     * work/home/fax/cell phone numbers, work/business/other
	     * addresses.
	     */
	    $phone = $row['Phone'.$row['PhoneDisplay']];
	    $phonelabel = $row['Phone'.$row['PhoneDisplay'].'Type'];
	    if ( !empty( $phone ) )
	      $phone = $phone . "(" . $this->phonelabelnames[$phonelabel] . ")";
	    $workphone = $homephone = $fax = $cellphone = "";
	    for ($i = 1; $i <= 5; $i++) {
		if (!empty($row['Phone'.$i])) {
		    switch ($row['Phone'.$i.'Type']) {
			case PHONE_LABEL_WORK:
			    $workphone = $row['Phone'.$i]; break;
			case PHONE_LABEL_HOME:
			    $homephone = $row['Phone'.$i]; break;
			case PHONE_LABEL_FAX:
			    $fax       = $row['Phone'.$i]; break;
			case PHONE_LABEL_MOBILE:
			    $cellphone = $row['Phone'.$i]; break;
		    }
		}
	    }
	    if ( !empty( $row['EmailSelector'] ) )
	      $email = $row['Email'.$row['EmailSelector']];
	    else
	      $email = $row['Email1'];
	    $homeaddress = trim(
		(!empty($row['HomeStreet']) ? "{$row['HomeStreet']}\n" : "") .
	 	(!empty($row['HomeCity']) ? "{$row['HomeCity']}, " : "") .
		$row['HomeState'] .
		(!empty($row['HomePostalCode']) ? " {$row['HomePostalCode']}" : ""));
	    $businessaddress = trim( $row['BusinessStreet'] .
		(!empty($row['BusinessStreet']) ? "\n" : "") . $row['BusinessCity'] .
		(!empty($row['BusinessCity']) ? ", " : "") . $row['BusinessState'] .
		(!empty($row['BusinessPostalCode']) ? " {$row['BusinessPostalCode']}" : ""));
	    $otheraddress = trim( $row['OtherStreet'] .
		(!empty($row['OtherStreet']) ? "\n" : "") . $row['OtherCity'] .
		(!empty($row['OtherCity']) ? ", " : "") . $row['OtherState'] .
		(!empty($row['OtherPostalCode']) ? " {$row['OtherPostalCode']}" : ""));

	    switch ( $row['AddressSelector'] ) {
		case 'B':
		    $mailingaddress = $businessaddress; break;
		case 'O':
		    $mailingaddress = $otheraddress;    break;
		case 'H':
		default:
		    $mailingaddress = $homeaddress;     break;
	    }
	    $ret = array('name'      => "{$row['FirstName']} {$row['LastName']}",
			 'phonedisplay' => $phone,
			 'homephone' => $homephone,
			 'workphone' => $workphone,
			 'cellphone' => $cellphone,
			 'fax'       => $fax,
                         'email'     => $email,
			 'homeaddress'    => $homeaddress,
			 'businessaddress' => $businessaddress,
			 'otheraddress'   => $otheraddress,
			 'mailingaddress' => $mailingaddress,
			 'modified'  => $row['Modified'],
			 'created'   => $row['Created'],
                         'backend'   => $this->bnum,
                         'source'    => &$this->sname);

	    /* Fill the return array with each of the SQL fields */
	    foreach( $this->fnames as $field => $sqlname )
		$ret[$field] = $row[$sqlname];
	    return ($ret);
        }
        return array();
    }

    /* List all addresses */
    function &list_addr() {
	global $active_cat;

        $ret = array();
        if (!$this->open()) {
            return false;
        }

        $query = sprintf("SELECT * FROM %s WHERE created_by='%s'",
                         $this->table, $this->owner);

	if (!empty( $active_cat ) > 0) {
	  if ($active_cat != 'Unfiled')
	      $query .= " AND Categories LIKE '%$active_cat%'";
	  else
	      $query .= " AND Categories = ''";
	}
	$query .= " ORDER BY $this->sortby";

        $res = $this->dbh->query($query);
        
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }

        while ($row = $res->fetchRow(DB_FETCHMODE_ASSOC)) {

	    /* Construct composite fields:  primary phone, primary email.
	     */
	    $phone = $row['Phone'.$row['PhoneDisplay']];
	    $phonelabel = $row['Phone'.$row['PhoneDisplay'].'Type'];
	    if ( !empty( $phone ) )
	      $phone = $phone . "(" . $this->phonelabelnames[$phonelabel] . ")";
	    if ( !empty( $row['EmailSelector'] ) )
	      $email = $row['Email'.$row['EmailSelector']];
	    else
	      $email = $row['Email1'];
            array_push($ret, array('uid'  => $row['uid'],
                                   'name'      => "{$row['FirstName']} {$row['LastName']}",
                                   'firstname' => $row['FirstName'],
                                   'lastname'  => $row['LastName'],
                                   'company'   => $row['Company'],
				   'jobtitle'  => $row['JobTitle'],
                                   'email'     => $email,
				   'phonedisplay' => $phone,
                                   'backend'   => $this->bnum,
                                   'source'    => &$this->sname));
        }
        return $ret;
    }

    /* Add address */
    function add($userdata) {

        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
        /* See if user exist already */
        $ret = $this->lookup($userdata[uid]);
        if (!empty($ret)) {
            return $this->set_error(sprintf(_("User '%s' already exist"),
                                            $ret['uid']));
        }

        /* Create query */
	$query = "INSERT INTO " . $this->table . "(created_by";
	$query2 = ") VALUES('" . $this->owner . "'";
	while (list($field, $sqlname) = each ($this->fnames)) {
	  $query .= ", " . $sqlname;
	  $query2 .= ",'" . $this->dbh->quoteString($userdata[$field]) . "'";
	}
	$query .= ",Created" . $query2 . ",NOW())";

         /* Do the insert */
         $r = $this->dbh->simpleQuery($query);
         if ($r == DB_OK) {
             return true;
         }

         /* Fail */
         return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Delete address */
    function remove($alias) {
        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("DELETE FROM %s WHERE created_by='%s' AND (",
                         $this->table, $this->owner);

        $sepstr = '';
        while (list($undef, $uid) = each($alias)) {
            $query .= sprintf("%s uid='%s' ", $sepstr,
                              $this->dbh->quoteString($uid));
            $sepstr = 'OR';

	    /* Log UID as deleted */
	    $r = $this->dbh->simpleQuery(
	             "INSERT INTO " . $this->history .
    		     " (uid,modby,field) VALUES ('$uid','$this->owner'," .
		     "'DELETED')");
        }
        $query .= ')';

        /* Delete entry */
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Modify address */
    function modify($alias, $userdata) {

        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
         /* See if user exist */
        $ret = $this->lookup($alias);
        if (empty($ret)) {
            return $this->set_error(sprintf(_("User '%s' does not exist"),
                                            $alias));
        }

	/* Log any changes */
	foreach( $this->fnames as $field => $sqlname) {
	    if ($ret[$field] != $userdata[$field] && $field != 'categories') {
		$query = "INSERT INTO " . $this->history .
		  " (uid,modby,field,newval,oldval) VALUES ('$userdata[uid]'," . 
		  "'$this->owner'," .
		  "'$field','" . $this->dbh->quoteString($userdata[$field]) .
		  "','" . $this->dbh->quoteString($ret[$field]) . "')";
		$r = $this->dbh->simpleQuery($query);
	    }
	}

        /* Create query */
	$query = "UPDATE " . $this->table . " SET ";

	$comma = '';
	foreach( $this->fnames as $field => $sqlname ) {
	  $query .= $comma . $sqlname .
		    "='" . $this->dbh->quoteString($userdata[$field]) . "'";
	  $comma = ", ";
	}		 
	$query .= " WHERE created_by='" . $this->owner . "' AND uid='" .
		 $this->dbh->quoteString($alias) . "'";

//echo $query . '<BR>';

        /* Do the insert */
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                        DB::errorMessage($r)));
    }

    /* Query server for category names */
    function category_get() {

	$query = "DESCRIBE " . $this->table . " Categories";
	
        $res = $this->dbh->query($query);
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	$row = $res->fetchRow(DB_FETCHMODE_ASSOC);
	$set = $row['Type'];
	$set = substr($set,5,strlen($set)-7); // Strip "set(" ... ");"
	return preg_split("/','/",$set); // Split into an array
    }


    /* Add a new category name */
    function category_create($category) {

	/* First get the existing categories */
	$query = "DESCRIBE " . $this->table . " Categories";
	
        $res = $this->dbh->query($query);
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	$row = $res->fetchRow(DB_FETCHMODE_ASSOC);
	$set = $row['Type'];

	/* Now tack on the new category */
	$query = "ALTER TABLE " . $this->table . " CHANGE Categories " .
	    " Categories SET(" . substr($set,4,strlen($set)-5) . ",'" .
	    $category . "') NOT NULL";
	
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Delete category names */
    function category_drop($catlist) {

	/* First get the existing categories */
	$query = "DESCRIBE " . $this->table . " Categories";
	
        $res = $this->dbh->query($query);
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	$row = $res->fetchRow(DB_FETCHMODE_ASSOC);
	$set = $row['Type'];

	$newcat = array_diff (explode("','", substr($set, 5, strlen($set)-7)),
			      $catlist);
	
	/* Now tack on the new category */
	$query = "ALTER TABLE " . $this->table . " CHANGE Categories " .
	    " Categories SET('" . implode("','", $newcat) . "') NOT NULL";
	
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Add records to a category */
    function category_addrecs($category, $uids) {
        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("UPDATE %s SET Categories=CONCAT_WS(',',`Categories`,'%s') WHERE created_by='%s' AND (",
                         $this->table, $category, $this->owner) .
                 "uid = '" . implode("' OR uid='", $uids) . "')";

        /* Invoke the entry */
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Add records from a category to another category */
    function category_move($category, $oldcategory) {
        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("UPDATE %s SET Categories=CONCAT_WS(',',`Categories`,'%s') WHERE created_by='%s' AND Categories LIKE '%%%s%%'",
                         $this->table, $category, $this->owner, $oldcategory);

        /* Invoke the entry */
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Remove records from a category */
    function category_rmrecs($category, $uids) {
        if (!$this->writeable) {
            return $this->set_error(_("Addressbook is read-only"));
        }

        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("UPDATE %s SET Categories=Categories&~pow(2,find_in_set('%s',Categories)-1) WHERE created_by='%s' AND (",
                         $this->table, $category, $this->owner) .
                 "uid = '" . implode("' OR uid='", $uids) . "')";

        /* Invoke the entry */
        $r = $this->dbh->simpleQuery($query);
        if ($r == DB_OK) {
            return true;
        }

        /* Fail */
        return $this->set_error(sprintf(_("Database error: %s"),
                                         DB::errorMessage($r)));
    }

    /* Count records in a category */
    function category_count($category) {

        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("SELECT COUNT(*) FROM %s WHERE Categories LIKE '%%%s%%' AND created_by='%s'",
                         $this->table, $category, $this->owner);

        /* Invoke the entry */
        $res = $this->dbh->query($query);
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	$row = $res->fetchRow(DB_FETCHMODE_ASSOC);
	return $row['COUNT(*)'];
    }

    /*
     * set_sort
     *   Parameter should look like 'field1 [ASC|DESC], field2 ...'.
     */
    function set_sort ($fields) {

	$params = explode( ",", $fields);
	$newsort = '';
	$comma = '';
	foreach ($params as $param) {
	    $field = preg_split('/ /', trim($param));
	    if (!isset ($this->fnames[strtolower($field[0])]))
		return false;
	    else {
		$newsort .= $comma . $this->fnames[strtolower($field[0])];
		if (isset($field[1]))
		   $newsort .= " $field[1]";
		$comma = ', ';
	    }
	}
        $this->sortby = $newsort;
        return true;
    }

    /*
     * get_history
     */
    function get_history ($uid) {
        if (!$this->open()) {
            return false;
        }
         
        /* Create query */
        $query = sprintf("SELECT * FROM %s WHERE uid='%s'",
                         $this->history, $uid);

        /* Invoke the entry */
        $res = $this->dbh->query($query);
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	$ret = array();
        while ($row = $res->fetchRow(DB_FETCHMODE_ASSOC)) {
            array_push($ret, array('moddate' => $row['moddate'],
				   'modby'   => $row['modby'],
                                   'field'   => $row['field'],
                                   'newval'  => $row['newval'],
                                   'oldval'  => $row['oldval']));
        }
        return $ret;
    }

    /*
     * raw_sql_cmd
     */
    function raw_sql_cmd ($query) {
        if (!$this->open()) {
            return false;
        }
         
        /* Invoke the entry */
        $res = $this->dbh->query(str_replace('INSERT INTO contacts ',
		'INSERT INTO '.$this->table.' ', $query));
        if (DB::isError($res)) {
            return $this->set_error(sprintf(_("Database error: %s"),
                                            DB::errorMessage($res)));
        }
	return true;
    }
} /* End of class abook_database */

?>
