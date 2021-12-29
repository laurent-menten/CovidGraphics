package be.lmenten.covid;

public interface Operation<T extends Number>
{
	T operation( T arg1, T arg2 );
}
