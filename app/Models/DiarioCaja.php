<?php
 
namespace App\Models;
 
use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;
 
class DiarioCaja extends Model
{
    protected $table = 'diario_caja';
    use HasFactory;
 
    protected $fillable = [
        'fecha',
        'venta_revo_iva_incluido',
        'caja_fuerte_inicio',
        'efectivo_diario',
        'tarjetas',
        'cover_manager',
        'transferencia',
        'propinas',
        'otras_formas_pago',
        'exceso_quebranto',
        'retiros',
        'bancos',
        'empresas_seguridad',
        'caja_fuerte_final'
    ];
}