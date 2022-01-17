/*
a executer dans php devel sur le site de l'iSm2, puis a import en csv format utf8 separateur points virgules
*/
header("content-type:application/csv;charset=UTF-8");
header("Content-Disposition:attachment;filename=\"permanents.csv\"");
echo "type;nom;prenom;mail;equipe;orcid;halid\n";
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
	$orcid=$account->field_user_orcis[LANGUAGE_NONE][0]['value'];
	/*$excludepubs = implode(',', array_column($account->field_user_exclude_pub[LANGUAGE_NONE], "value"));*/
	echo "$type;$user->nom;$user->prenom;$account->mail;$equipe->name;$orcid;$listhalid\n";
}