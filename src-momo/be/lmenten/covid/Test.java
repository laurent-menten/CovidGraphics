package be.lmenten.covid;

public class Test {

	public static void main( String[] args )
	{
		DayOfYear<Double> doy1 = new DayOfYear<>( 1, 1, new Double [5] );
		doy1.setData( 0, 0.d );
		doy1.setData( 1, 1.d );
		doy1.setData( 2, 2.d );
		doy1.setData( 3, 3.d );
		doy1.updateData( 4, 4.d, (a,b) -> { return (a == null) ? b : a + b; } );
		doy1.compute();

		System.out.println( "min: " + doy1.getMin() );
		System.out.println( "max: " + doy1.getMax() );
		System.out.println( "avg: " + doy1.getAverage() );
		System.out.println( "std: " + doy1.getStandartDeviation() );
		
		DayOfYear<Integer> doy2 = new DayOfYear<>( 1, 1, new Integer [5] );
		doy2.setData( 0, 0 );
		doy2.setData( 1, 1 );
		doy2.setData( 2, 2 );
		doy2.setData( 3, 3 );
		doy2.updateData( 4, 4, (a,b) -> { return (a == null) ? b : a + b; } );
		doy2.compute();

		System.out.println( "min: " + doy2.getMin() );
		System.out.println( "max: " + doy2.getMax() );
		System.out.println( "avg: " + doy2.getAverage() );
		System.out.println( "std: " + doy2.getStandartDeviation() );
}

}
