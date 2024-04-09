<?php
 
namespace Database\Seeders;
 
use Illuminate\Database\Seeder;
use App\Models\DiarioCaja;
use Faker\Factory as Faker;
 
class DiarioCajaSeeder extends Seeder
{
    public function run()
    {
        // Creamos una instancia de Faker
        $faker = Faker::create();
 
        // Creamos un solo registro de ejemplo
        DiarioCaja::create([
            'fecha' => $faker->date(),
            'venta_revo_iva_incluido' => $faker->randomFloat(2, 0, 10000),
            'caja_fuerte_inicio' => $faker->randomFloat(2, 0, 10000),
            'efectivo_diario' => $faker->randomFloat(2, 0, 10000),
            'tarjetas' => $faker->randomFloat(2, 0, 10000),
            'cover_manager' => $faker->word,
            'transferencia' => $faker->word,
            'propinas' => $faker->randomFloat(2, 0, 10000),
            'otras_formas_pago' => $faker->word,
            'exceso_quebranto' => $faker->randomFloat(2, 0, 10000),
            'retiros' => $faker->randomFloat(2, 0, 10000),
            'bancos' => $faker->word,
            'empresas_seguridad' => $faker->randomFloat(2, 0, 10000),
            'caja_fuerte_final' => $faker->randomFloat(2, 0, 10000),
        ]);
    }
}
?>