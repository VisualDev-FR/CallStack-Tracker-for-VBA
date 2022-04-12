
    function initDossierDeplier()
    {
        var oArbo = document.getElementById('historique'),
            aDossier = oArbo.getElementsByTagName('input');
        
        for(var i = 0; i <aDossier.length; i++)
        {
            aDossier[i].checked=false;
        }
        
        aDossier[0].checked=true;
    }

    document.addEventListener('DOMContentLoaded',function()
    {
        initDossierDeplier();
    });
