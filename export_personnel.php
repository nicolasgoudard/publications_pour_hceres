header("Content-Disposition:attachment;filename=\"permanents.csv\"");
echo "type;nom;prenom;mail;equipe;orcid;halid;actif;date entree;sorti;date sortie\n";
$accounts =entity_load('user');
foreach ($accounts as $account) {
	$user=_get_user_from_account($account);
    if (isUserOfType($user,"permanent"))
    	$type="permanent";
    else {
        $term_type = taxonomy_term_load($account->field_user_fonction[LANGUAGE_NONE][0]['tid']);
        $type=$term_type->name;
    }
	$equipe = taxonomy_term_load($user->id_equipe);
	$listhalid = implode(',', $user->halId);
	$orcid=$account->field_user_orcid[LANGUAGE_NONE][0]['value'];
        $actif=$account->status;
        if ($account->field_user_date_entree[LANGUAGE_NONE][0]['value'] != "")
	        $date_entree=date("d/m/Y", strtotime($account->field_user_date_entree[LANGUAGE_NONE][0]['value']));
	    else
	        $date_entree="";
        $sorti=$account->field_user_sorti[LANGUAGE_NONE][0]['value'];
        if ($account->field_user_date_sortie[LANGUAGE_NONE][0]['value'] != "")
       		$date_sortie=date("d/m/Y", strtotime( $account->field_user_date_sortie[LANGUAGE_NONE][0]['value']));
       	else
       		$date_sortie="";
	/*$excludepubs = implode(',', array_column($account->field_user_exclude_pub[LANGUAGE_NONE], "value"));*/
	echo "$type;$user->nom;$user->prenom;$account->mail;$equipe->name;$orcid;$listhalid;$actif;$date_entree;$sorti;$date_sortie;\n";
}