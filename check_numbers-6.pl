#  menor ou igual anterior
#  separar o último dígito 
#  conferir se a sessão bate o número
#  

use utf8::all;          # Turn on UTF-8 on everything
if ($^O == "MSWin32"){  # Fixing outupt in CMD and PowerShell
    system("chcp 65001"); 
}
print $^O, "\n";

##############################################

use strict;
use warnings;
use Spreadsheet::XLSX;
use Scalar::Util qw(looks_like_number);


our $previous_prontu = 0;
our $previous_sess   = 0;

my $excel = Spreadsheet::XLSX -> new ('data.xlsx');

# criar vars para nº da sessão para comparar com o val


foreach my $sheet (@{$excel -> {Worksheet}}){
  printf("Sheet: %s\n", $sheet->{Name});
  
  foreach my $row ($sheet -> {MinRow} .. $sheet -> {MaxRow}){
     my $session  =  $sheet -> {Cells} [$row] [0];
     my $prontu   =  $sheet -> {Cells} [$row] [1];
     
     if ($prontu) {
        my $prontu_content = $prontu  -> {Val};
        my $sess_content = 5555;
        $sess_content = $session -> {Val} if ($session);

        CheckSession($row, $prontu_content, $sess_content);

        unless (looks_like_number($prontu_content)){
            PrintCurrentCell($row, $prontu_content, $sess_content, "  → → → prontuário não é nº");
            $previous_prontu = 00000000;
            goto FAZ_ALGO_MELHOR_DEPOIS;
        }

        if (sprintf("%s",$sess_content) eq sprintf("%s", $previous_sess)){
            if ($prontu_content < $previous_prontu){
                my $custom = sprintf("  → → → %d menor que: %d", $prontu_content, $previous_prontu);
                PrintCurrentCell($row, $prontu_content, $sess_content, $custom); 
            }
        }
        
        if ($prontu_content == $previous_prontu){
            my $custom = sprintf( "  → → →  %d = a último: %d", $prontu_content, $previous_prontu);
            PrintCurrentCell($row, $prontu_content, $sess_content, $custom); 
        }
        
        $previous_prontu = $prontu_content;
        
        FAZ_ALGO_MELHOR_DEPOIS:
        $previous_sess = $sess_content;
    }
  }
}


print "\n\n";
exit;

sub CheckSession {
    my ($row, $prontu, $sess) = @_;
    my $prontu_tmp = 0;
    my $extract_sess = 0;

    #tira último número
    unless (!looks_like_number($prontu)) {$prontu_tmp = ($prontu/10)} 
    else {goto GAMBIARRA}

    $prontu_tmp = sprintf("%d", $prontu_tmp); #tira o float
    $extract_sess = (substr($prontu_tmp, -2)); #mantém somente os 2 últimos nº

    #checa se tem zero à esquerda
    unless (!looks_like_number($extract_sess) or
            !looks_like_number($sess) or
            length($extract_sess) < 2) {
            if (substr($extract_sess, -2, 1) == 0) { $extract_sess = substr($extract_sess, -1) } }
    else { goto GAMBIARRA }


    GAMBIARRA:
    if (!looks_like_number($sess)){
        PrintCurrentCell($row, $prontu, $sess, " → → → sessão não é número");
    }
    elsif ($sess != $extract_sess){
        PrintCurrentCell($row, $prontu, $sess, " → → → sessão não bate");
    }
}

sub PrintCurrentCell {
    my ($row, $prontu,$sess, $custom) = @_;
    unless (!looks_like_number($sess)) {$sess = "VAZIA" if ($sess == 5555)}
    print "\n\n|ROW: ", $row, " |SESS: ", $sess, " |PRONT: ", $prontu, $custom;
}




#open(OUTPUT_FILE, ">", "results.txt") or die $!;
#print OUTPUT_FILE "Line 1\n";
#print OUTPUT_FILE "Line 2\n";
#close(OUTPUT_FILE);

#Encode::_utf8_on($prontu_content);
