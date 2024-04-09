<?php
 
use Illuminate\Database\Migrations\Migration;
use Illuminate\Database\Schema\Blueprint;
use Illuminate\Support\Facades\Schema;
 
return new class extends Migration
{
    /**
     * Run the migrations.
     */
    public function up(): void
    {
            Schema::create('diario_caja', function (Blueprint $table) {
                $table->id();
                $table->date('fecha');
                $table->decimal('venta_revo_iva_incluido', 10, 2);
                $table->decimal('caja_fuerte_inicio', 10, 2);
                $table->decimal('efectivo_diario', 10, 2);
                $table->decimal('tarjetas', 10, 2);
                $table->string('cover_manager');
                $table->string('transferencia');
                $table->decimal('propinas');
                $table->string('otras_formas_pago');
                $table->decimal('exceso_quebranto', 10, 2);
                $table->decimal('retiros', 10, 2);
                $table->string('bancos');
                $table->decimal('empresas_seguridad', 10, 2);
                $table->decimal('caja_fuerte_final', 10, 2);
                $table->timestamps();
        });
    }
 
    /**
     * Reverse the migrations.
     */
    public function down(): void
    {
        Schema::dropIfExists('diario_caja');
    }
};