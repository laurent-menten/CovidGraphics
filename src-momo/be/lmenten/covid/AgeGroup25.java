package be.lmenten.covid;

public enum AgeGroup25
{
	AGE_0_24		( "0-24",   0, 24 ),
	AGE_25_44		( "25-44", 25, 44 ),
	AGE_45_64		( "45-64", 45, 64 ),
	AGE_65_74		( "65-74", 65, 74 ),
	AGE_75_84		( "75-84", 75, 84 ),
	AGE_85_PLUS		( "85+", 85 ),

	UNKNOWN			( "?" )
	;

	private final String id;
	private final int base;
	private final int limit;

	private AgeGroup25( String id )
	{
		this( id, -1, -1 );
	}

	private AgeGroup25( String id, int base )
	{
		this( id, base, -1 );
	}

	private AgeGroup25( String id, int base, int limit )
	{
		this.id = id;
		this.base = base;
		this.limit = limit;
	}

	public static AgeGroup25 lookup( String id )
	{
		if( id != null )
		{
			if( id.startsWith( "'" ) )
			{
				id = id.substring( 1 );
			}

			for( AgeGroup25 ageClass : AgeGroup25.values() )
			{
				if( ageClass.id.equals( id ) )
				{
					return ageClass;
				}
			}
		}

		return UNKNOWN;
	}

	public String getId()
	{
		return id;
	}

	public int getBase()
	{
		return base;
	}

	public int getLimit()
	{
		return limit;
	}

	public int getIndex()
	{
		return ordinal();
	}

	public AgeGroup25 lookup( int age )
	{
		if( age < 0 )
		{
			throw new IndexOutOfBoundsException( age );
		}

		for( AgeGroup25 c : AgeGroup25.values() )
		{
			if( age >= c.base && (c.limit != -1 && age <= c.limit) )
			{
				return c;
			}
		}

		return UNKNOWN;
	}
}
